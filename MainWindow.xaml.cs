using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Media;
using Microsoft.Win32;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Windows;
using ClosedXML.Excel;

namespace CadFilesUpdater.Windows
{
    public partial class MainWindow : System.Windows.Window
    {
        private sealed class FileEntry : System.ComponentModel.INotifyPropertyChanged
        {
            private string _filePath;
            private bool _isSelected;
            private int _blocksWithAttributes;

            public string FilePath
            {
                get => _filePath;
                set
                {
                    if (_filePath == value) return;
                    _filePath = value;
                    OnPropertyChanged(nameof(FilePath));
                }
            }

            public bool IsSelected
            {
                get => _isSelected;
                set
                {
                    if (_isSelected == value) return;
                    _isSelected = value;
                    OnPropertyChanged(nameof(IsSelected));
                }
            }

            public int BlocksWithAttributes
            {
                get => _blocksWithAttributes;
                set
                {
                    if (_blocksWithAttributes == value) return;
                    _blocksWithAttributes = value;
                    OnPropertyChanged(nameof(BlocksWithAttributes));
                }
            }

            public event System.ComponentModel.PropertyChangedEventHandler PropertyChanged;

            private void OnPropertyChanged(string propertyName)
            {
                PropertyChanged?.Invoke(this, new System.ComponentModel.PropertyChangedEventArgs(propertyName));
            }
        }

        private sealed class BlockEntry
        {
            public string BlockName { get; set; }
            public bool IsSelected { get; set; }
        }

        private sealed class LastEditedContext
        {
            public string BlockName { get; set; }
            public string AttributeTag { get; set; }
            public string Value { get; set; }
        }

        private readonly ObservableCollection<FileEntry> _files = new ObservableCollection<FileEntry>();
        private readonly ObservableCollection<BlockEntry> _blocks = new ObservableCollection<BlockEntry>();

        // Per file: offline-scanned instances (cached so toggling selections doesn't re-scan).
        private readonly Dictionary<string, List<BlockAnalyzer.AttributeInstanceRow>> _instancesByFile =
            new Dictionary<string, List<BlockAnalyzer.AttributeInstanceRow>>(StringComparer.OrdinalIgnoreCase);

        // Cache of edits: (filePath, blockHandle, blockName, tag) -> value
        private readonly Dictionary<BlockAnalyzer.ChangeKey, string> _changes =
            new Dictionary<BlockAnalyzer.ChangeKey, string>();

        // Column display order preservation (column name -> DisplayIndex)
        private Dictionary<string, int> _savedColumnOrder = null;

        // Undo/Redo system
        private sealed class ChangeSnapshot
        {
            public Dictionary<BlockAnalyzer.ChangeKey, string> Changes { get; set; }
            public string Description { get; set; }
            public List<string> RemovedFiles { get; set; } // Files that were removed in this action
        }
        private readonly List<ChangeSnapshot> _undoHistory = new List<ChangeSnapshot>();
        private int _undoHistoryIndex = -1;
        private const int MaxUndoHistory = 50;

        private LastEditedContext _lastEdited;
        // Remember last values used in Apply window (when no table edit was made)
        private string _lastApplyValue = "";
        private string _lastApplyBlockName = "";
        private string _lastApplyAttributeTag = "";
        private DataTable _table;
        // Track file order for alternating colors in non-editable columns
        private readonly List<string> _fileOrder = new List<string>();
        private int? _lastBlockClickIndex;
        private bool? _lastBlockRangeState;
        private bool _suppressSelectionEvents;
        private System.Windows.Threading.DispatcherTimer _refreshTimer;
        private bool _refreshBlocks;
        private bool _refreshGrid;
        private System.Windows.Threading.DispatcherTimer _gridStyleRefreshTimer;
        private bool _allFilesSelected = true;
        private bool _allBlocksSelected = true;

        public MainWindow()
        {
            InitializeComponent();
            FilesListView.ItemsSource = _files;
            BlocksListView.ItemsSource = _blocks;
            InitDebouncedRefresh();
            RebuildGrid();
            UpdateSummaries();
            
            // Setup keyboard shortcuts for undo/redo
            KeyDown += MainWindow_KeyDown;
            
            UpdateUndoRedoButtons();
        }

        private void InitDebouncedRefresh()
        {
            _refreshTimer = new System.Windows.Threading.DispatcherTimer
            {
                Interval = TimeSpan.FromMilliseconds(120)
            };
            _refreshTimer.Tick += (s, e) =>
            {
                _refreshTimer.Stop();
                if (_refreshBlocks) RebuildBlocksList();
                if (_refreshGrid) RebuildGrid();
                UpdateSummaries();
                _refreshBlocks = false;
                _refreshGrid = false;
            };
        }

        private void ScheduleGridStyleRefresh()
        {
            if (_gridStyleRefreshTimer == null)
            {
                _gridStyleRefreshTimer = new System.Windows.Threading.DispatcherTimer
                {
                    Interval = TimeSpan.FromMilliseconds(60)
                };
                _gridStyleRefreshTimer.Tick += (s, e) =>
                {
                    _gridStyleRefreshTimer.Stop();
                    RefreshGridCellStylesForVisibleRows();
                };
            }
            _gridStyleRefreshTimer.Stop();
            _gridStyleRefreshTimer.Start();
        }

        private void RefreshGridCellStylesForVisibleRows()
        {
            try
            {
                for (int i = 0; i < AttributesDataGrid.Items.Count; i++)
                {
                    var row = (DataGridRow)AttributesDataGrid.ItemContainerGenerator.ContainerFromIndex(i);
                    if (row != null)
                        RefreshGridCellStylesForRow(row);
                }
            }
            catch
            {
                // best-effort
            }
        }

        private void ScheduleRefresh(bool blocks, bool grid)
        {
            _refreshBlocks |= blocks;
            _refreshGrid |= grid;
            _refreshTimer.Stop();
            _refreshTimer.Start();
        }

        private void SetBusy(bool isBusy, string text)
        {
            try
            {
                BusyText.Text = string.IsNullOrWhiteSpace(text) ? "Working..." : text;
                BusyOverlay.Visibility = isBusy ? Visibility.Visible : Visibility.Collapsed;
                // Force a render pass so the overlay appears before heavy work starts.
                Dispatcher.Invoke(() => { }, System.Windows.Threading.DispatcherPriority.Render);
            }
            catch
            {
                // Best-effort only.
            }
        }

        private void SetBusyProgress(int processed, int total)
        {
            try
            {
                if (total <= 0)
                {
                    BusyProgress.IsIndeterminate = true;
                    return;
                }

                BusyProgress.IsIndeterminate = false;
                BusyProgress.Minimum = 0;
                BusyProgress.Maximum = total;
                BusyProgress.Value = processed;
            }
            catch
            {
                // Best-effort only.
            }
        }

        private void AddFiles_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new Microsoft.Win32.OpenFileDialog
            {
                Multiselect = true,
                Filter = "AutoCAD files (*.dwg)|*.dwg|All files (*.*)|*.*",
                Title = "Select AutoCAD files"
            };

            if (dialog.ShowDialog() != true) return;

            AddFilesFromPaths(dialog.FileNames);
        }

        private void AddOptions_Click(object sender, RoutedEventArgs e)
        {
            var btn = sender as Button;
            if (btn?.ContextMenu == null) return;
            btn.ContextMenu.PlacementTarget = btn;
            btn.ContextMenu.IsOpen = true;
        }

        private void AddDrawingsMenu_Click(object sender, RoutedEventArgs e)
        {
            AddFiles_Click(sender, e);
        }

        private void AddFolders_Click(object sender, RoutedEventArgs e)
        {
            var includeSubfolders = MessageBox.Show(
                "Do you want to include subfolders?",
                "Add folders",
                MessageBoxButton.YesNoCancel,
                MessageBoxImage.Question);

            if (includeSubfolders == MessageBoxResult.Cancel) return;

            using (var dialog = new System.Windows.Forms.FolderBrowserDialog())
            {
                dialog.Description = "Select folder with DWG files";
                dialog.ShowNewFolderButton = false;
                if (dialog.ShowDialog() != System.Windows.Forms.DialogResult.OK) return;

                var option = includeSubfolders == MessageBoxResult.Yes
                    ? SearchOption.AllDirectories
                    : SearchOption.TopDirectoryOnly;

                var files = Directory.EnumerateFiles(dialog.SelectedPath, "*.dwg", option).ToList();
                if (files.Count == 0)
                {
                    MessageBox.Show("No DWG files found in the selected folder.",
                        "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                AddFilesFromPaths(files);
            }
        }

        private void AddOpenedDrawings_Click(object sender, RoutedEventArgs e)
        {
            var docManager = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager;
            var files = new List<string>();
            foreach (Document d in docManager)
            {
                if (!string.IsNullOrWhiteSpace(d.Name) && File.Exists(d.Name))
                    files.Add(d.Name);
            }

            if (files.Count == 0)
            {
                MessageBox.Show("No opened drawings with a valid file path were found.",
                    "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            AddFilesFromPaths(files);
        }

        private void AddFilesFromPaths(IEnumerable<string> filePaths)
        {
            var existing = new HashSet<string>(_files.Select(f => f.FilePath), StringComparer.OrdinalIgnoreCase);
            var addedAny = false;
            var newlyAdded = new List<FileEntry>();
            foreach (var fp in filePaths)
            {
                if (string.IsNullOrWhiteSpace(fp) || existing.Contains(fp)) continue;
                var entry = new FileEntry { FilePath = fp, IsSelected = false, BlocksWithAttributes = 0 };
                _files.Add(entry);
                newlyAdded.Add(entry);
                addedAny = true;
            }

            if (!addedAny)
            {
                UpdateSummaries();
                return;
            }

            try
            {
                SetBusy(true, "Analyzing files...");
                EnsureInstancesForSelectedFiles();
                RebuildBlocksList();
                RebuildGrid();
                UpdateSummaries();
            }
            finally
            {
                SetBusy(false, null);
            }

            if (!_allFilesSelected && newlyAdded.Count > 0)
            {
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    foreach (var entry in newlyAdded)
                        FilesListView.SelectedItems.Add(entry);
                }), System.Windows.Threading.DispatcherPriority.Loaded);
            }
        }

        private void RemoveFiles_Click(object sender, RoutedEventArgs e)
        {
            var toRemove = _allFilesSelected
                ? _files.ToList()
                : FilesListView.SelectedItems.Cast<FileEntry>().ToList();
            
            if (toRemove.Count == 0) return;

            EnsureUndoBaseline();

            foreach (var fe in toRemove)
            {
                _files.Remove(fe);
                _instancesByFile.Remove(fe.FilePath);
            }

            // Drop changes for removed files
            var removedPaths = new HashSet<string>(toRemove.Select(t => t.FilePath), StringComparer.OrdinalIgnoreCase);
            var keysToRemove = _changes.Keys.Where(k => removedPaths.Contains(k.FilePath)).ToList();
            foreach (var k in keysToRemove) _changes.Remove(k);

            // Record undo state after removing
            RecordUndoState("Remove files", removedFiles: toRemove.Select(f => f.FilePath).ToList());

            RebuildBlocksList();
            RebuildGrid();
            UpdateSummaries();
            UpdateUndoRedoButtons();

            if (FilesListView.SelectedItems.Count == 0)
            {
                _allFilesSelected = true;
                RebuildBlocksList();
                RebuildGrid();
                UpdateSummaries();
            }
        }

        private void AllFiles_Click(object sender, RoutedEventArgs e)
        {
            _allFilesSelected = true;
            FilesListView.SelectedItems.Clear();
            EnsureInstancesForSelectedFiles();
            RebuildBlocksList();
            RebuildGrid();
            UpdateSummaries();
        }

        private void FilesListView_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (FilesListView.SelectedItems.Count > 0)
                _allFilesSelected = false;

            ScheduleRefresh(blocks: true, grid: true);
        }

        // Selection is handled by the DataGrid; we only react in SelectionChanged.

        private void BlocksListView_PreviewMouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (IsClickInsideCheckBox(e.OriginalSource as DependencyObject)) return;

            var lvi = FindAncestor<ListViewItem>(e.OriginalSource as DependencyObject);
            if (lvi == null) return;

            var entry = lvi.DataContext as BlockEntry;
            if (entry == null) return;

            var idx = BlocksListView.ItemContainerGenerator.IndexFromContainer(lvi);
            var isShift = (System.Windows.Input.Keyboard.Modifiers & System.Windows.Input.ModifierKeys.Shift) == System.Windows.Input.ModifierKeys.Shift;

            if (isShift && _lastBlockClickIndex.HasValue && idx >= 0)
            {
                int start = Math.Min(_lastBlockClickIndex.Value, idx);
                int end = Math.Max(_lastBlockClickIndex.Value, idx);
                var desired = _lastBlockRangeState ?? true;
                for (int i = start; i <= end && i < _blocks.Count; i++)
                    _blocks[i].IsSelected = desired; // shift-click = select/unselect range based on anchor
            }
            else
            {
                entry.IsSelected = !entry.IsSelected;
                _lastBlockRangeState = entry.IsSelected;
            }

            if (idx >= 0) _lastBlockClickIndex = idx;
            BlocksListView.Items.Refresh();
            ScheduleRefresh(blocks: false, grid: true);
            e.Handled = true;
        }

        private void BlocksListView_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (BlocksListView.SelectedItems.Count > 0)
                _allBlocksSelected = false;

            ScheduleRefresh(blocks: false, grid: true);
        }

        private void AllBlocks_Click(object sender, RoutedEventArgs e)
        {
            _allBlocksSelected = true;
            BlocksListView.SelectedItems.Clear();
            ScheduleRefresh(blocks: true, grid: true);
        }

        private static bool IsClickInsideCheckBox(DependencyObject d)
        {
            return FindAncestor<CheckBox>(d) != null;
        }

        private static T FindAncestor<T>(DependencyObject d) where T : DependencyObject
        {
            while (d != null)
            {
                if (d is T t) return t;
                d = VisualTreeHelper.GetParent(d);
            }
            return null;
        }

        private void BlocksSelectionChanged(object sender, RoutedEventArgs e)
        {
            if (_suppressSelectionEvents) return;

            // Support Shift-range when clicking the checkbox too.
            var cb = sender as CheckBox;
            if (cb != null)
            {
                var lvi = ItemsControl.ContainerFromElement(BlocksListView, cb) as ListViewItem;
                var idx = lvi != null ? BlocksListView.ItemContainerGenerator.IndexFromContainer(lvi) : -1;
                var isShift = (System.Windows.Input.Keyboard.Modifiers & System.Windows.Input.ModifierKeys.Shift) == System.Windows.Input.ModifierKeys.Shift;
                var entry = lvi != null ? (lvi.DataContext as BlockEntry) : null;
                var currentState = entry != null && entry.IsSelected;

                if (isShift && _lastBlockClickIndex.HasValue && _lastBlockRangeState.HasValue && idx >= 0)
                {
                    int start = Math.Min(_lastBlockClickIndex.Value, idx);
                    int end = Math.Max(_lastBlockClickIndex.Value, idx);
                    try
                    {
                        _suppressSelectionEvents = true;
                        for (int i = start; i <= end && i < _blocks.Count; i++)
                            _blocks[i].IsSelected = _lastBlockRangeState.Value;
                        BlocksListView.Items.Refresh();
                    }
                    finally
                    {
                        _suppressSelectionEvents = false;
                    }
                }
                if (idx >= 0)
                {
                    if (!isShift)
                        _lastBlockRangeState = currentState;
                    _lastBlockClickIndex = idx;
                }
            }

            ScheduleRefresh(blocks: false, grid: true);
        }

        private void SelectAllBlocks_Click(object sender, RoutedEventArgs e)
        {
            foreach (var b in _blocks) b.IsSelected = true;
            BlocksListView.Items.Refresh();
            ScheduleRefresh(blocks: false, grid: true);
        }

        private void UnselectAllBlocks_Click(object sender, RoutedEventArgs e)
        {
            foreach (var b in _blocks) b.IsSelected = false;
            BlocksListView.Items.Refresh();
            ScheduleRefresh(blocks: false, grid: true);
        }

        private void ReloadFiles_Click(object sender, RoutedEventArgs e)
        {
            var targets = _files.Select(f => f.FilePath).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
            if (targets.Count == 0)
            {
                MessageBox.Show(this, "No files loaded.", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            try
            {
                SetBusyProgress(0, targets.Count);
                SetBusy(true, "Reloading files...");
                int processed = 0;
                foreach (var fp in targets)
                {
                    processed++;
                    BusyText.Text = $"Reloading {processed}/{targets.Count}: {System.IO.Path.GetFileNameWithoutExtension(fp)}";
                    SetBusyProgress(processed, targets.Count);
                    EnsureInstances(fp, forceReload: true);
                    Dispatcher.Invoke(() => { }, System.Windows.Threading.DispatcherPriority.Render);
                }

                // Preserve edits (_changes) and just refresh view.
                RebuildBlocksList();
                RebuildGrid();
                UpdateSummaries();
            }
            finally
            {
                SetBusy(false, null);
                BusyProgress.IsIndeterminate = true;
            }
        }

        private void AttributesDataGrid_AutoGeneratingColumn(object sender, DataGridAutoGeneratingColumnEventArgs e)
        {
            // DataView sometimes yields a pseudo-column (.) that just shows "System.Data.DataRowView".
            // Hide it.
            if (string.IsNullOrWhiteSpace(e.PropertyName) || e.PropertyName == "." ||
                string.Equals(e.PropertyName, "Row", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(e.PropertyName, "RowState", StringComparison.OrdinalIgnoreCase))
            {
                e.Cancel = true;
                return;
            }

            // Hide internal columns
            if (string.Equals(e.PropertyName, "FilePath", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(e.PropertyName, "Handle", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(e.PropertyName, "_FileGroupIndex", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(e.PropertyName, "_FileRowIndex", StringComparison.OrdinalIgnoreCase))
            {
                e.Cancel = true;
                return;
            }

            // Always set SortMemberPath so we can reliably map a column back to the DataTable column name
            // (headers are user-friendly and not stable identifiers).
            e.Column.SortMemberPath = e.PropertyName;

            // Friendly headers for non-editable columns (colors set dynamically in RefreshGridCellStylesForRow)
            if (string.Equals(e.PropertyName, "File", StringComparison.OrdinalIgnoreCase))
            {
                e.Column.Header = "Drawing";
                e.Column.IsReadOnly = true;
            }
            else if (string.Equals(e.PropertyName, "LayoutOwner", StringComparison.OrdinalIgnoreCase))
            {
                e.Column.Header = "Layout owner";
                e.Column.IsReadOnly = true;
            }
            else if (string.Equals(e.PropertyName, "BlockName", StringComparison.OrdinalIgnoreCase))
            {
                e.Column.Header = "Block name";
                e.Column.IsReadOnly = true;
            }
            else
            {
                // Attribute columns: set header and ensure value is visible
                e.Column.Header = e.PropertyName; // attribute tag
                
                if (e.Column is DataGridTextColumn textColumn)
                {
                    // Set ElementStyle to ensure value is always visible (even when selected)
                    var elementStyle = new Style(typeof(TextBlock));
                    elementStyle.Setters.Add(new Setter(TextBlock.PaddingProperty, new Thickness(4, 2, 4, 2)));
                    elementStyle.Setters.Add(new Setter(TextBlock.VerticalAlignmentProperty, VerticalAlignment.Center));
                    textColumn.ElementStyle = elementStyle;
                    
                    // Set EditingElementStyle to show value during editing
                    var editingStyle = new Style(typeof(TextBox));
                    editingStyle.Setters.Add(new Setter(TextBox.BorderThicknessProperty, new Thickness(0)));
                    editingStyle.Setters.Add(new Setter(TextBox.PaddingProperty, new Thickness(4, 2, 4, 2)));
                    textColumn.EditingElementStyle = editingStyle;
                }
            }
        }

        private void AttributesDataGrid_BeginningEdit(object sender, DataGridBeginningEditEventArgs e)
        {
            var drv = e.Row != null ? (e.Row.Item as DataRowView) : null;
            if (drv == null) return;

            var colName = e.Column != null ? e.Column.SortMemberPath : null;
            if (string.IsNullOrWhiteSpace(colName)) return;

            // No edits for N/A cells
            try
            {
                var current = drv.Row[colName]?.ToString();
                if (string.Equals(current, "N/A", StringComparison.OrdinalIgnoreCase))
                {
                    e.Cancel = true;
                    return;
                }
            }
            catch
            {
                // ignore
            }
        }

        private void AttributesDataGrid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            if (e.EditAction != DataGridEditAction.Commit) return;
            var drv = e.Row != null ? (e.Row.Item as DataRowView) : null;
            if (drv == null) return;

            var colName = e.Column != null ? e.Column.SortMemberPath : null;
            if (string.IsNullOrWhiteSpace(colName)) return;

            // Only attribute columns are editable
            if (colName.Equals("File", StringComparison.OrdinalIgnoreCase) ||
                colName.Equals("LayoutOwner", StringComparison.OrdinalIgnoreCase) ||
                colName.Equals("BlockName", StringComparison.OrdinalIgnoreCase))
                return;

            var filePath = drv.Row["FilePath"]?.ToString();
            var handle = drv.Row["Handle"]?.ToString();
            var blockName = drv.Row["BlockName"]?.ToString();
            if (string.IsNullOrWhiteSpace(filePath) || string.IsNullOrWhiteSpace(handle) || string.IsNullOrWhiteSpace(blockName))
                return;

            var editor = e.EditingElement as TextBox;
            var newValue = editor?.Text ?? "";

            var key = new BlockAnalyzer.ChangeKey(filePath, handle, blockName, colName);
            // If user typed the same value as original, don't keep it as a "changed" cell.
            var original = GetOriginalAttributeValue(filePath, handle, colName);
            var isSameAsOriginal = string.Equals(original ?? "", newValue ?? "", StringComparison.Ordinal);
            var hadChange = _changes.TryGetValue(key, out var existingValue);
            var willChange = isSameAsOriginal
                ? hadChange
                : (!hadChange || !string.Equals(existingValue ?? "", newValue ?? "", StringComparison.Ordinal));

            if (willChange)
                EnsureUndoBaseline();

            if (isSameAsOriginal)
                _changes.Remove(key);
            else
                _changes[key] = newValue;

            _lastEdited = new LastEditedContext
            {
                BlockName = blockName,
                AttributeTag = colName,
                Value = newValue
            };
            
            // Record undo state for single cell edit
            if (willChange)
                RecordUndoState("Edit cell");

            RefreshGridCellStylesForRow(e.Row);
        }

        private void AttributesDataGrid_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            RefreshGridCellStylesForRow(e.Row);
        }

        private string GetOriginalAttributeValue(string filePath, string handle, string tag)
        {
            EnsureInstances(filePath);
            if (!_instancesByFile.TryGetValue(filePath, out var rows)) return null;
            for (int i = 0; i < rows.Count; i++)
            {
                var r = rows[i];
                if (!string.Equals(r.BlockHandle, handle, StringComparison.OrdinalIgnoreCase)) continue;
                if (r.Attributes.TryGetValue(tag, out var v)) return v ?? "";
                return null;
            }
            return null;
        }

        private void RefreshAllGridCellStyles()
        {
            // Defer until visuals exist. Use Render priority to ensure cells are created.
            Dispatcher.BeginInvoke(new Action(() =>
            {
                // First pass: refresh visible rows immediately
                RefreshGridCellStylesForVisibleRows();
                
                // Second pass: ensure all rows get refreshed after a short delay
                // This handles cases where cells are created lazily due to virtualization
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    RefreshGridCellStylesForVisibleRows();
                }), System.Windows.Threading.DispatcherPriority.Loaded);
            }), System.Windows.Threading.DispatcherPriority.Render);
        }

        private void RefreshGridCellStylesForRow(DataGridRow row)
        {
            try
            {
                if (row == null) return;
                var drv = row.Item as DataRowView;
                if (drv == null) return;

                var filePath = drv.Row["FilePath"]?.ToString();
                var handle = drv.Row["Handle"]?.ToString();
                var blockName = drv.Row["BlockName"]?.ToString();
                if (string.IsNullOrWhiteSpace(filePath) || string.IsNullOrWhiteSpace(handle) || string.IsNullOrWhiteSpace(blockName))
                    return;

                // Ensure row template is applied and cells are generated
                if (!row.IsLoaded)
                {
                    row.UpdateLayout();
                }

                int fileGroupIndex = 0;
                int fileRowIndex = 0;
                try
                {
                    if (drv.Row.Table.Columns.Contains("_FileGroupIndex"))
                        fileGroupIndex = Convert.ToInt32(drv.Row["_FileGroupIndex"]);
                    if (drv.Row.Table.Columns.Contains("_FileRowIndex"))
                        fileRowIndex = Convert.ToInt32(drv.Row["_FileRowIndex"]);
                }
                catch
                {
                    // best-effort
                }

                bool isEvenFile = (fileGroupIndex % 2 == 0);
                bool isEvenRecord = (fileRowIndex % 2 == 0);

                // File grouping colors (yellow/blue) with alternating shades per record
                var yellowLight = Color.FromRgb(0xFE, 0xF9, 0xC3); // existing light yellow
                var yellowDark = Color.FromRgb(0xF9, 0xF3, 0xA8);  // existing darker yellow
                var orangeLight = Color.FromRgb(0xFD, 0xE7, 0xD0); // light orange
                var orangeDark = Color.FromRgb(0xF9, 0xC8, 0x9B);  // darker orange
                var editableLight = Colors.White;
                var editableDark = Color.FromRgb(0xF3, 0xF4, 0xF6); // light gray

                Color nonEditableBgColor = isEvenFile
                    ? (isEvenRecord ? yellowLight : yellowDark)
                    : (isEvenRecord ? orangeLight : orangeDark);
                Color editableBgColor = isEvenRecord ? editableLight : editableDark;

                for (int colIndex = 0; colIndex < AttributesDataGrid.Columns.Count; colIndex++)
                {
                    var col = AttributesDataGrid.Columns[colIndex];
                    var colName = col.SortMemberPath ?? (col.Header?.ToString() ?? "");
                    
                    var cell = GetCell(row, colIndex);
                    if (cell == null)
                    {
                        // If cell doesn't exist yet, try to force generation
                        row.UpdateLayout();
                        cell = GetCell(row, colIndex);
                        if (cell == null) continue;
                    }
                    
                    // Set background for non-editable columns (File, LayoutOwner, BlockName) - yellow/blue by file + shade by record
                    if (colName == "File" || colName == "LayoutOwner" || colName == "BlockName")
                    {
                        cell.Background = new SolidColorBrush(nonEditableBgColor);
                        cell.Foreground = Brushes.Black;
                        cell.ClearValue(DataGridCell.FontStyleProperty);
                        continue;
                    }

                    // Editable attribute columns - always white (unless changed or N/A)
                    string cellValue = null;
                    try { cellValue = drv.Row[colName]?.ToString(); } catch { }

                    var key = new BlockAnalyzer.ChangeKey(filePath, handle, blockName, colName);
                    if (_changes.ContainsKey(key))
                    {
                        // Changed cell: light green
                        cell.Background = new SolidColorBrush(Color.FromRgb(0xBB, 0xF7, 0xD0)); // light green
                        cell.Foreground = Brushes.Black;
                        cell.FontStyle = FontStyles.Normal;
                    }
                    else if (string.Equals(cellValue, "N/A", StringComparison.OrdinalIgnoreCase))
                    {
                        // N/A cell: keep non-editable background (yellow/blue + shade)
                        cell.Background = new SolidColorBrush(nonEditableBgColor);
                        cell.Foreground = Brushes.Gray;
                        cell.FontStyle = FontStyles.Italic;
                    }
                    else
                    {
                        // Normal editable cell: white or light gray (by record)
                        cell.Background = new SolidColorBrush(editableBgColor);
                        cell.Foreground = Brushes.Black;
                        cell.ClearValue(DataGridCell.FontStyleProperty);
                    }
                }
            }
            catch
            {
                // best-effort
            }
        }

        private DataGridCell GetCell(DataGridRow row, int columnIndex)
        {
            if (row == null) return null;

            DataGridCellsPresenter presenter = FindVisualChild<DataGridCellsPresenter>(row);
            if (presenter == null)
            {
                row.ApplyTemplate();
                presenter = FindVisualChild<DataGridCellsPresenter>(row);
            }
            if (presenter == null) return null;

            var cell = presenter.ItemContainerGenerator.ContainerFromIndex(columnIndex) as DataGridCell;
            return cell;
        }

        private void AttributesDataGrid_ScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            // Both horizontal and vertical scrolling create/destroy cells because of virtualization.
            // Refresh styles for currently realized cells so colors stay correct while scrolling.
            if (Math.Abs(e.HorizontalChange) > 0.0 ||
                Math.Abs(e.VerticalChange) > 0.0 ||
                Math.Abs(e.ExtentWidthChange) > 0.0 ||
                Math.Abs(e.ExtentHeightChange) > 0.0 ||
                Math.Abs(e.ViewportWidthChange) > 0.0 ||
                Math.Abs(e.ViewportHeightChange) > 0.0)
            {
                ScheduleGridStyleRefresh();
            }
        }

        private void AttributesDataGrid_Sorting(object sender, DataGridSortingEventArgs e)
        {
            // Let the default sorting happen, then recompute row indices for alternating shades.
            Dispatcher.BeginInvoke(new Action(() =>
            {
                UpdateFileRowIndicesFromView();
                RefreshGridCellStylesForVisibleRows();
            }), System.Windows.Threading.DispatcherPriority.Loaded);
        }

        private static T FindVisualChild<T>(DependencyObject obj) where T : DependencyObject
        {
            if (obj == null) return null;
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(obj); i++)
            {
                var child = VisualTreeHelper.GetChild(obj, i);
                if (child is T t) return t;
                var result = FindVisualChild<T>(child);
                if (result != null) return result;
            }
            return null;
        }

        private void ApplySimilar_Click(object sender, RoutedEventArgs e)
        {
            var allFiles = _files.Select(f => f.FilePath).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
            var defaultSelected = GetSelectedFilePaths();
            
            // Get current values from selected cell or last change, not from _lastEdited (which may be outdated after undo)
            string initialValue = "";
            string initialBlockName = "";
            string initialAttributeTag = "";
            
            // Try to get from currently selected cell
            var selectedCell = AttributesDataGrid.CurrentCell;
            if (selectedCell.Item is DataRowView drv && selectedCell.Column != null)
            {
                var colName = selectedCell.Column.SortMemberPath;
                if (!string.IsNullOrWhiteSpace(colName) &&
                    !colName.Equals("File", StringComparison.OrdinalIgnoreCase) &&
                    !colName.Equals("LayoutOwner", StringComparison.OrdinalIgnoreCase) &&
                    !colName.Equals("BlockName", StringComparison.OrdinalIgnoreCase))
                {
                    var filePath = drv.Row["FilePath"]?.ToString();
                    var handle = drv.Row["Handle"]?.ToString();
                    var blockName = drv.Row["BlockName"]?.ToString();
                    if (!string.IsNullOrWhiteSpace(filePath) && !string.IsNullOrWhiteSpace(handle) && !string.IsNullOrWhiteSpace(blockName))
                    {
                        var key = new BlockAnalyzer.ChangeKey(filePath, handle, blockName, colName);
                        if (_changes.TryGetValue(key, out var currentValue))
                        {
                            initialValue = currentValue;
                            initialBlockName = blockName;
                            initialAttributeTag = colName;
                        }
                    }
                }
            }
            
            // If no selected cell with change, try to find last change from _changes
            if (string.IsNullOrWhiteSpace(initialAttributeTag) && _changes.Count > 0)
            {
                var lastChange = _changes.Last();
                initialValue = lastChange.Value;
                initialBlockName = lastChange.Key.BlockName;
                initialAttributeTag = lastChange.Key.AttributeTag;
            }
            
            // Fallback to remembered values
            if (string.IsNullOrWhiteSpace(initialAttributeTag))
            {
                initialValue = _lastApplyValue ?? "";
                initialBlockName = _lastApplyBlockName ?? "";
                initialAttributeTag = _lastApplyAttributeTag ?? "";
            }
            
            var dlg = new ApplySimilarWindow(
                initialValue,
                initialBlockName,
                initialAttributeTag,
                allFiles,
                defaultSelected,
                fp =>
                {
                    EnsureInstances(fp);
                    return _instancesByFile.TryGetValue(fp, out var rows) ? rows : new List<BlockAnalyzer.AttributeInstanceRow>();
                })
            { Owner = this };
            if (dlg.ShowDialog() != true) return;

            var newValue = dlg.ValueText ?? "";
            var mode = dlg.Mode;
            
            // Remember values used in Apply window (for next time, if no table edit)
            _lastApplyValue = newValue;
            _lastApplyBlockName = dlg.SelectedBlockName ?? "";
            _lastApplyAttributeTag = dlg.SelectedAttributeTag ?? "";

            var selectedScope = dlg.SelectedFilePathsForApply != null && dlg.SelectedFilePathsForApply.Count > 0
                ? dlg.SelectedFilePathsForApply
                : GetSelectedFilePaths();
            var targetFiles = mode == ApplySimilarMode.SelectedFiles
                ? selectedScope
                : _files.Select(f => f.FilePath).ToList();

            // Apply changes (including empty value if field is empty)
            bool anyChange = false;
            bool baselineEnsured = false;
            foreach (var fp in targetFiles)
            {
                EnsureInstances(fp);
                if (!_instancesByFile.TryGetValue(fp, out var rows)) continue;
                foreach (var row in rows)
                {
                    if (!row.BlockName.Equals(dlg.SelectedBlockName, StringComparison.OrdinalIgnoreCase)) continue;
                    if (!row.Attributes.ContainsKey(dlg.SelectedAttributeTag)) continue;
                    var key = new BlockAnalyzer.ChangeKey(fp, row.BlockHandle, row.BlockName, dlg.SelectedAttributeTag);
                    // Always set value, even if empty (to clear the attribute)
                    var hadChange = _changes.TryGetValue(key, out var existingValue);
                    var willChange = !hadChange || !string.Equals(existingValue ?? "", newValue ?? "", StringComparison.Ordinal);
                    if (willChange)
                    {
                        if (!baselineEnsured)
                        {
                            EnsureUndoBaseline();
                            baselineEnsured = true;
                        }
                        _changes[key] = newValue;
                        anyChange = true;
                    }
                }
            }
            
            // Record undo state for bulk edit
            if (anyChange)
                RecordUndoState("Bulk edit: Apply to all attributes");

            RebuildGrid();
            UpdateSummaries();
        }

        private void SaveChanges_Click(object sender, RoutedEventArgs e)
        {
            if (_changes.Count == 0)
            {
                MessageBox.Show(this, "There are no changes to save.", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var filePaths = _files.Select(f => f.FilePath).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
            UpdateResult result;
            try
            {
                SetBusyProgress(0, filePaths.Count);
                SetBusy(true, "Saving changes...");
                result = BlockAnalyzer.SaveCachedChanges(_changes, filePaths,
                    (processed, total, file) =>
                    {
                        GridSummaryText.Text = $"Saving {processed}/{total}: {System.IO.Path.GetFileNameWithoutExtension(file)}";
                        BusyText.Text = $"Saving {processed}/{total}: {System.IO.Path.GetFileNameWithoutExtension(file)}";
                        SetBusyProgress(processed, total);
                        Dispatcher.Invoke(() => { }, System.Windows.Threading.DispatcherPriority.Render);
                    });
            }
            finally
            {
                SetBusy(false, null);
                BusyProgress.IsIndeterminate = true;
            }

            var msg = $"Processed: {result.TotalFiles}\nSuccessful: {result.SuccessfulFiles}\nFailed: {result.FailedFiles}";
            if (result.Errors.Count > 0)
            {
                msg += "\n\nErrors:\n" + string.Join("\n", result.Errors.Select(e2 => $"{System.IO.Path.GetFileNameWithoutExtension(e2.FilePath)}: {e2.ErrorMessage}"));
            }

            // Don't activate window - let user stay in Cursor/other app
            MessageBox.Show(this, msg, "Save result", MessageBoxButton.OK,
                result.FailedFiles == 0 ? MessageBoxImage.Information : MessageBoxImage.Warning);

            // After save: refresh instance cache for selected files (values may have changed).
            EnsureInstancesForSelectedFiles(forceReload: true);
            RebuildBlocksList();
            RebuildGrid();
            UpdateSummaries();
        }

        private void EnsureInstancesForSelectedFiles(bool forceReload = false)
        {
            foreach (var fp in GetSelectedFilePaths())
            {
                EnsureInstances(fp, forceReload);
            }
        }

        private void EnsureInstances(string filePath, bool forceReload = false)
        {
            if (!forceReload && _instancesByFile.ContainsKey(filePath)) return;
            try
            {
                var rows = BlockAnalyzer.AnalyzeAttributeInstances(filePath);
                _instancesByFile[filePath] = rows;
                var entry = _files.FirstOrDefault(f => string.Equals(f.FilePath, filePath, StringComparison.OrdinalIgnoreCase));
                if (entry != null)
                    entry.BlocksWithAttributes = rows?.Count ?? 0;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[MainWindow] AnalyzeAttributeInstances error: {filePath}: {ex.Message}");
                _instancesByFile[filePath] = new List<BlockAnalyzer.AttributeInstanceRow>();
                var entry = _files.FirstOrDefault(f => string.Equals(f.FilePath, filePath, StringComparison.OrdinalIgnoreCase));
                if (entry != null)
                    entry.BlocksWithAttributes = 0;
            }
        }

        private List<string> GetSelectedFilePaths()
        {
            if (_allFilesSelected)
                return _files.Select(f => f.FilePath).ToList();

            return FilesListView.SelectedItems.Cast<FileEntry>()
                .Select(f => f.FilePath)
                .ToList();
        }

        private HashSet<string> GetSelectedBlocks()
        {
            if (_allBlocksSelected)
                return _blocks.Select(b => b.BlockName).ToHashSet(StringComparer.OrdinalIgnoreCase);

            return BlocksListView.SelectedItems.Cast<BlockEntry>()
                .Select(b => b.BlockName)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
        }

        private void RebuildBlocksList()
        {
            var selectedFiles = GetSelectedFilePaths();
            var blockNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var fp in selectedFiles)
            {
                EnsureInstances(fp);
                if (!_instancesByFile.TryGetValue(fp, out var rows)) continue;
                foreach (var r in rows) blockNames.Add(r.BlockName);
            }

            _blocks.Clear();
            foreach (var bn in blockNames.OrderBy(x => x))
            {
                _blocks.Add(new BlockEntry { BlockName = bn, IsSelected = false });
            }
            BlocksListView.Items.Refresh();
        }

        private void RebuildGrid()
        {
            var selectedFiles = GetSelectedFilePaths();
            var selectedBlocks = GetSelectedBlocks();

            var visibleRows = new List<BlockAnalyzer.AttributeInstanceRow>();
            if (selectedBlocks.Count == 0)
            {
                // Requirement: if no blocks are selected, show nothing.
                selectedFiles.Clear(); // avoids scanning
            }
            
            // Track file order for alternating colors
            _fileOrder.Clear();
            var seenFiles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            
            foreach (var fp in selectedFiles)
            {
                EnsureInstances(fp);
                if (!_instancesByFile.TryGetValue(fp, out var rows)) continue;
                foreach (var r in rows)
                {
                    if (!selectedBlocks.Contains(r.BlockName)) continue;
                    visibleRows.Add(r);
                    
                    // Track unique files in order they appear
                    if (!seenFiles.Contains(r.FilePath))
                    {
                        seenFiles.Add(r.FilePath);
                        _fileOrder.Add(r.FilePath);
                    }
                }
            }

            // Sort rows: first by file, then by layout owner, then by block name
            visibleRows = visibleRows.OrderBy(r => r.FilePath, StringComparer.OrdinalIgnoreCase)
                                    .ThenBy(r => r.LayoutName, StringComparer.OrdinalIgnoreCase)
                                    .ThenBy(r => r.BlockName, StringComparer.OrdinalIgnoreCase)
                                    .ToList();

            var attributeTags = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var r in visibleRows)
                foreach (var t in r.Attributes.Keys)
                    attributeTags.Add(t);

            _table = new DataTable();
            _table.Columns.Add("File", typeof(string));
            _table.Columns.Add("LayoutOwner", typeof(string));
            _table.Columns.Add("BlockName", typeof(string));
            _table.Columns.Add("FilePath", typeof(string)); // hidden
            _table.Columns.Add("Handle", typeof(string));   // hidden
            _table.Columns.Add("_FileGroupIndex", typeof(int)); // hidden
            _table.Columns.Add("_FileRowIndex", typeof(int));   // hidden

            foreach (var tag in attributeTags.OrderBy(x => x))
                _table.Columns.Add(tag, typeof(string));

            var fileRowCounters = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            foreach (var r in visibleRows)
            {
                var dr = _table.NewRow();
                var fileGroupIndex = _fileOrder.IndexOf(r.FilePath);
                if (fileGroupIndex < 0) fileGroupIndex = 0;
                var fileRowIndex = fileRowCounters.TryGetValue(r.FilePath, out var existingIndex) ? existingIndex : 0;
                fileRowCounters[r.FilePath] = fileRowIndex + 1;
                dr["File"] = System.IO.Path.GetFileNameWithoutExtension(r.FilePath);
                dr["LayoutOwner"] = r.LayoutName;
                dr["BlockName"] = r.BlockName;
                dr["FilePath"] = r.FilePath;
                dr["Handle"] = r.BlockHandle;
                dr["_FileGroupIndex"] = fileGroupIndex;
                dr["_FileRowIndex"] = fileRowIndex;

                foreach (var tag in attributeTags)
                {
                    var baseVal = r.Attributes.TryGetValue(tag, out var v) ? v : "N/A";
                    var key = new BlockAnalyzer.ChangeKey(r.FilePath, r.BlockHandle, r.BlockName, tag);
                    if (_changes.TryGetValue(key, out var changed))
                        dr[tag] = changed;
                    else
                        dr[tag] = baseVal;
                }

                _table.Rows.Add(dr);
            }

            AttributesDataGrid.ItemsSource = _table.DefaultView;
            GridSummaryText.Text = $"{visibleRows.Count} row(s)";
            
            // Restore column display order if it was saved
            RestoreColumnDisplayOrder();
            
            UpdateFileRowIndicesFromView();
            RefreshAllGridCellStyles();
        }

        private void SaveColumnDisplayOrder()
        {
            _savedColumnOrder = new Dictionary<string, int>();
            foreach (var col in AttributesDataGrid.Columns)
            {
                var colName = col.SortMemberPath ?? (col.Header?.ToString() ?? "");
                if (!string.IsNullOrEmpty(colName))
                {
                    _savedColumnOrder[colName] = col.DisplayIndex;
                }
            }
        }

        private void RestoreColumnDisplayOrder()
        {
            if (_savedColumnOrder == null || _savedColumnOrder.Count == 0)
                return;

            // Wait for columns to be generated - use Render priority to ensure columns exist
            Dispatcher.BeginInvoke(new Action(() =>
            {
                try
                {
                    if (AttributesDataGrid.Columns.Count == 0)
                        return;

                    // Build a list of columns with their saved DisplayIndex
                    var columnsToReorder = new List<(DataGridColumn col, int displayIndex)>();
                    
                    foreach (var col in AttributesDataGrid.Columns)
                    {
                        var colName = col.SortMemberPath ?? (col.Header?.ToString() ?? "");
                        if (!string.IsNullOrEmpty(colName) && _savedColumnOrder.TryGetValue(colName, out int savedIndex))
                        {
                            columnsToReorder.Add((col, savedIndex));
                        }
                    }

                    if (columnsToReorder.Count == 0)
                        return;

                    // Sort by saved DisplayIndex and assign new DisplayIndex sequentially
                    // This preserves relative order while handling removed/added columns
                    var sorted = columnsToReorder.OrderBy(x => x.displayIndex).ToList();
                    for (int i = 0; i < sorted.Count; i++)
                    {
                        sorted[i].col.DisplayIndex = i;
                    }
                }
                catch
                {
                    // best-effort, ignore errors
                }
            }), System.Windows.Threading.DispatcherPriority.Render);
        }

        private void UpdateSummaries()
        {
            var selectedCount = _allFilesSelected
                ? _files.Count
                : FilesListView.SelectedItems.Count;
            FilesSummaryText.Text = $"{_files.Count} file(s) loaded, {selectedCount} selected";
        }

        private void MainWindow_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if ((System.Windows.Input.Keyboard.Modifiers & System.Windows.Input.ModifierKeys.Control) == System.Windows.Input.ModifierKeys.Control)
            {
                if (e.Key == System.Windows.Input.Key.C)
                {
                    if (ShouldHandleGridClipboard())
                    {
                        if (CopySelectionToClipboard())
                            e.Handled = true;
                        return;
                    }
                }
                else if (e.Key == System.Windows.Input.Key.V)
                {
                    if (ShouldHandleGridClipboard())
                    {
                        if (TryPasteClipboardToSelection())
                            e.Handled = true;
                        return;
                    }
                }
            }

            if (e.Key == System.Windows.Input.Key.Delete)
            {
                if (ShouldHandleGridClipboard())
                {
                    if (TryClearSelectedCells())
                        e.Handled = true;
                    return;
                }
            }

            if (e.Key == System.Windows.Input.Key.Z)
            {
                if ((System.Windows.Input.Keyboard.Modifiers & System.Windows.Input.ModifierKeys.Control) == System.Windows.Input.ModifierKeys.Control)
                {
                    if ((System.Windows.Input.Keyboard.Modifiers & System.Windows.Input.ModifierKeys.Shift) == System.Windows.Input.ModifierKeys.Shift)
                    {
                        // Ctrl+Shift+Z = Redo
                        if (RedoButton.IsEnabled)
                        {
                            Redo_Click(null, null);
                            e.Handled = true;
                        }
                    }
                    else
                    {
                        // Ctrl+Z = Undo
                        if (UndoButton.IsEnabled)
                        {
                            Undo_Click(null, null);
                            e.Handled = true;
                        }
                    }
                }
            }
        }

        private void AttributesDataGrid_PreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Delete)
            {
                if (TryClearSelectedCells())
                {
                    e.Handled = true;
                }
            }
        }

        #region Undo/Redo

        private void EnsureUndoBaseline()
        {
            if (_undoHistory.Count == 0)
                RecordUndoState("Initial state");
        }

        private void RecordUndoState(string description, List<string> removedFiles = null)
        {
            // Remove any history after current index (when user did undo and then made new change)
            if (_undoHistoryIndex < _undoHistory.Count - 1)
                _undoHistory.RemoveRange(_undoHistoryIndex + 1, _undoHistory.Count - _undoHistoryIndex - 1);

            // Create deep copy snapshot of current changes
            var snapshot = new ChangeSnapshot
            {
                Changes = new Dictionary<BlockAnalyzer.ChangeKey, string>(),
                Description = description,
                RemovedFiles = removedFiles != null ? new List<string>(removedFiles) : null
            };
            foreach (var kv in _changes)
                snapshot.Changes[kv.Key] = kv.Value;

            _undoHistory.Add(snapshot);
            _undoHistoryIndex = _undoHistory.Count - 1;

            // Limit history size
            if (_undoHistory.Count > MaxUndoHistory)
            {
                _undoHistory.RemoveAt(0);
                _undoHistoryIndex--;
            }

            UpdateUndoRedoButtons();
        }

        private void Undo_Click(object sender, RoutedEventArgs e)
        {
            if (_undoHistoryIndex < 0) return;

            // Get current snapshot before undoing
            var currentSnapshot = _undoHistory[_undoHistoryIndex];
            
            _undoHistoryIndex--;
            if (_undoHistoryIndex >= 0)
            {
                var snapshot = _undoHistory[_undoHistoryIndex];
                _changes.Clear();
                foreach (var kv in snapshot.Changes)
                    _changes[kv.Key] = kv.Value;
                
                // Restore removed files if any
                if (currentSnapshot.RemovedFiles != null && currentSnapshot.RemovedFiles.Count > 0)
                {
                    foreach (var filePath in currentSnapshot.RemovedFiles)
                    {
                        // Check if file was already restored
                        if (!_files.Any(f => string.Equals(f.FilePath, filePath, StringComparison.OrdinalIgnoreCase)))
                        {
                            _files.Add(new FileEntry { FilePath = filePath, IsSelected = false });
                            // Reload instances for restored file
                            EnsureInstances(filePath);
                        }
                    }
                }
            }
            else
            {
                _changes.Clear();
                
                // Restore removed files if any
                if (currentSnapshot.RemovedFiles != null && currentSnapshot.RemovedFiles.Count > 0)
                {
                    foreach (var filePath in currentSnapshot.RemovedFiles)
                    {
                        if (!_files.Any(f => string.Equals(f.FilePath, filePath, StringComparison.OrdinalIgnoreCase)))
                        {
                            _files.Add(new FileEntry { FilePath = filePath, IsSelected = false });
                            EnsureInstances(filePath);
                        }
                    }
                }
            }

            RebuildBlocksList();
            RebuildGrid();
            UpdateSummaries();
            UpdateUndoRedoButtons();
        }

        private void Redo_Click(object sender, RoutedEventArgs e)
        {
            if (_undoHistoryIndex >= _undoHistory.Count - 1) return;

            _undoHistoryIndex++;
            var snapshot = _undoHistory[_undoHistoryIndex];
            _changes.Clear();
            foreach (var kv in snapshot.Changes)
                _changes[kv.Key] = kv.Value;
            
            // Remove files that were removed in this action
            if (snapshot.RemovedFiles != null && snapshot.RemovedFiles.Count > 0)
            {
                var toRemove = _files.Where(f => snapshot.RemovedFiles.Contains(f.FilePath, StringComparer.OrdinalIgnoreCase)).ToList();
                foreach (var fe in toRemove)
                {
                    _files.Remove(fe);
                    _instancesByFile.Remove(fe.FilePath);
                }
                
                // Drop changes for removed files
                var removedPaths = new HashSet<string>(snapshot.RemovedFiles, StringComparer.OrdinalIgnoreCase);
                var keysToRemove = _changes.Keys.Where(k => removedPaths.Contains(k.FilePath)).ToList();
                foreach (var k in keysToRemove) _changes.Remove(k);
            }

            RebuildBlocksList();
            RebuildGrid();
            UpdateSummaries();
            UpdateUndoRedoButtons();
        }

        private void UpdateUndoRedoButtons()
        {
            UndoButton.IsEnabled = _undoHistoryIndex >= 0;
            RedoButton.IsEnabled = _undoHistoryIndex < _undoHistory.Count - 1;
        }

        #endregion

        #region Context Menu

        private void AttributesDataGrid_ContextMenuOpening(object sender, RoutedEventArgs e)
        {
            // Show/hide menu items based on whether the clicked cell is editable
            var cell = GetSingleTargetCell();
            bool isEditable = false;
            bool hasCell = false;
            
            if (cell.HasValue)
            {
                var cellValue = cell.Value;
                var colName = cellValue.Column?.SortMemberPath;
                if (!string.IsNullOrWhiteSpace(colName))
                {
                    hasCell = true;
                    isEditable = IsAttributeColumn(colName);
                }
            }
            
            // Find menu items by name (they're defined in XAML with x:Name)
            var contextMenu = sender as System.Windows.Controls.ContextMenu;
            if (contextMenu != null)
            {
                foreach (var item in contextMenu.Items)
                {
                    if (item is System.Windows.Controls.MenuItem menuItem)
                    {
                        var name = menuItem.Name;
                        if (name == "ContextMenu_Copy")
                        {
                            menuItem.Visibility = hasCell ? System.Windows.Visibility.Visible : System.Windows.Visibility.Collapsed;
                        }
                        else if (name == "ContextMenu_RestoreCell" || 
                                 name == "ContextMenu_Paste" || 
                                 name == "ContextMenu_Delete" ||
                                 name == "ContextMenu_ApplySimilar")
                        {
                            menuItem.Visibility = isEditable ? System.Windows.Visibility.Visible : System.Windows.Visibility.Collapsed;
                        }
                    }
                }
            }
        }

        private void ContextMenu_CopyCell_Click(object sender, RoutedEventArgs e)
        {
            CopySelectionToClipboard();
        }

        private void ContextMenu_PasteCell_Click(object sender, RoutedEventArgs e)
        {
            TryPasteClipboardToSelection();
        }

        private void ContextMenu_DeleteCell_Click(object sender, RoutedEventArgs e)
        {
            TryClearSelectedCells();
        }

        private void ContextMenu_ApplySimilar_Click(object sender, RoutedEventArgs e)
        {
            // Get the value from the clicked cell
            var cell = GetSingleTargetCell();
            if (!cell.HasValue) return;

            var cellValue = cell.Value;
            var drv = cellValue.Item as DataRowView;
            var colName = cellValue.Column?.SortMemberPath;
            if (drv == null || string.IsNullOrWhiteSpace(colName)) return;
            if (!IsAttributeColumn(colName)) return;

            var filePath = drv.Row["FilePath"]?.ToString();
            var handle = drv.Row["Handle"]?.ToString();
            var blockName = drv.Row["BlockName"]?.ToString();
            if (string.IsNullOrWhiteSpace(filePath) || string.IsNullOrWhiteSpace(handle) || string.IsNullOrWhiteSpace(blockName))
                return;

            // Get the value from _changes if it exists, otherwise from the row
            string valueToApply = "";
            var sourceKey = new BlockAnalyzer.ChangeKey(filePath, handle, blockName, colName);
            if (_changes.TryGetValue(sourceKey, out var changedValue))
            {
                valueToApply = changedValue ?? "";
            }
            else
            {
                valueToApply = drv.Row[colName]?.ToString() ?? "";
            }

            // If the value is "N/A", don't apply it
            if (string.Equals(valueToApply, "N/A", StringComparison.OrdinalIgnoreCase))
                return;

            // Find all instances of this block across all files and apply the value
            int appliedCount = 0;
            bool anyChange = false;
            bool baselineEnsured = false;
            foreach (var fileInstances in _instancesByFile)
            {
                foreach (var row in fileInstances.Value)
                {
                    // Match by block name (case-insensitive)
                    if (string.Equals(row.BlockName, blockName, StringComparison.OrdinalIgnoreCase))
                    {
                        // Check if this block has this attribute
                        if (row.Attributes.ContainsKey(colName))
                        {
                            var key = new BlockAnalyzer.ChangeKey(row.FilePath, row.BlockHandle, row.BlockName, colName);
                            
                            // Get original value for this specific instance
                            var originalValue = GetOriginalAttributeValue(row.FilePath, row.BlockHandle, colName);
                            
                            // Only add to _changes if the value is different from original
                            if (string.Equals(originalValue ?? "", valueToApply ?? "", StringComparison.Ordinal))
                            {
                                // Value is same as original, remove from _changes if it exists
                                if (_changes.ContainsKey(key))
                                {
                                    if (!baselineEnsured)
                                    {
                                        EnsureUndoBaseline();
                                        baselineEnsured = true;
                                    }
                                    _changes.Remove(key);
                                    anyChange = true;
                                }
                            }
                            else
                            {
                                // Value is different, add to _changes
                                var hadChange = _changes.TryGetValue(key, out var existingValue);
                                var willChange = !hadChange || !string.Equals(existingValue ?? "", valueToApply ?? "", StringComparison.Ordinal);
                                if (willChange)
                                {
                                    if (!baselineEnsured)
                                    {
                                        EnsureUndoBaseline();
                                        baselineEnsured = true;
                                    }
                                    _changes[key] = valueToApply;
                                    appliedCount++;
                                    anyChange = true;
                                }
                            }
                        }
                    }
                }
            }

            if (anyChange)
                RecordUndoState("Apply value to all similar attributes");

            RebuildGridPreservingScroll();
            UpdateUndoRedoButtons();

            if (appliedCount > 0)
            {
                MessageBox.Show($"Applied value to {appliedCount} attribute(s) in block '{blockName}', attribute '{colName}' across all files.", 
                    "Apply Similar", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void ContextMenu_RevertAttribute_Click(object sender, RoutedEventArgs e)
        {
            var selectedCells = AttributesDataGrid.SelectedCells;
            if (selectedCells.Count == 0) return;

            var cellsToRevert = new HashSet<BlockAnalyzer.ChangeKey>();
            foreach (var cell in selectedCells)
            {
                var drv = cell.Item as DataRowView;
                if (drv == null) continue;

                var filePath = drv.Row["FilePath"]?.ToString();
                var handle = drv.Row["Handle"]?.ToString();
                var blockName = drv.Row["BlockName"]?.ToString();
                var colName = cell.Column?.SortMemberPath;

                if (string.IsNullOrWhiteSpace(filePath) || string.IsNullOrWhiteSpace(handle) ||
                    string.IsNullOrWhiteSpace(blockName) || string.IsNullOrWhiteSpace(colName))
                    continue;

                // Skip non-attribute columns
                if (colName.Equals("File", StringComparison.OrdinalIgnoreCase) ||
                    colName.Equals("LayoutOwner", StringComparison.OrdinalIgnoreCase) ||
                    colName.Equals("BlockName", StringComparison.OrdinalIgnoreCase))
                    continue;

                cellsToRevert.Add(new BlockAnalyzer.ChangeKey(filePath, handle, blockName, colName));
            }

            bool anyChange = false;
            foreach (var key in cellsToRevert)
            {
                if (_changes.ContainsKey(key))
                {
                    if (!anyChange) EnsureUndoBaseline();
                    _changes.Remove(key);
                    anyChange = true;
                }
            }

            if (anyChange)
                RecordUndoState("Revert attribute(s)");

            RebuildGridPreservingScroll();
            UpdateUndoRedoButtons();
        }

        private void ContextMenu_RevertBlock_Click(object sender, RoutedEventArgs e)
        {
            // Get cells from selection or current cell if nothing selected
            var cells = AttributesDataGrid.SelectedCells;
            if (cells.Count == 0 && AttributesDataGrid.CurrentCell.Item == null) return;

            // Collect block names to revert (across ALL files, not just the file from selected cell)
            var blockNamesToRevert = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            
            // Process selected cells
            foreach (var cellInfo in cells)
            {
                var drv = cellInfo.Item as DataRowView;
                if (drv == null) continue;

                var blockName = drv.Row["BlockName"]?.ToString();
                if (!string.IsNullOrWhiteSpace(blockName))
                    blockNamesToRevert.Add(blockName);
            }
            
            // If no selection but current cell exists, process it
            if (cells.Count == 0 && AttributesDataGrid.CurrentCell.Item != null)
            {
                var drv = AttributesDataGrid.CurrentCell.Item as DataRowView;
                if (drv != null)
                {
                    var blockName = drv.Row["BlockName"]?.ToString();
                    if (!string.IsNullOrWhiteSpace(blockName))
                        blockNamesToRevert.Add(blockName);
                }
            }

            // Remove all changes for these block names across ALL files (ignoring handle and filePath)
            var keysToRemove = _changes.Keys
                .Where(k => blockNamesToRevert.Contains(k.BlockName))
                .ToList();

            if (keysToRemove.Count > 0)
            {
                EnsureUndoBaseline();
                foreach (var key in keysToRemove)
                    _changes.Remove(key);
                RecordUndoState("Revert block(s)");
            }

            RebuildGridPreservingScroll();
            UpdateUndoRedoButtons();
        }

        private void ContextMenu_RevertFile_Click(object sender, RoutedEventArgs e)
        {
            var selectedCells = AttributesDataGrid.SelectedCells;
            if (selectedCells.Count == 0) return;

            var filesToRevert = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var cell in selectedCells)
            {
                var drv = cell.Item as DataRowView;
                if (drv == null) continue;

                var filePath = drv.Row["FilePath"]?.ToString();
                if (!string.IsNullOrWhiteSpace(filePath))
                    filesToRevert.Add(filePath);
            }

            var keysToRemove = _changes.Keys
                .Where(k => filesToRevert.Contains(k.FilePath))
                .ToList();

            if (keysToRemove.Count > 0)
            {
                EnsureUndoBaseline();
                foreach (var key in keysToRemove)
                    _changes.Remove(key);
                RecordUndoState("Revert file(s)");
            }

            RebuildGridPreservingScroll();
            UpdateUndoRedoButtons();
        }

        private void ContextMenu_OpenInAutoCAD_Click(object sender, RoutedEventArgs e)
        {
            var selectedCells = AttributesDataGrid.SelectedCells;
            if (selectedCells.Count == 0 && AttributesDataGrid.CurrentCell.Item == null) return;

            string filePath = null;
            if (selectedCells.Count > 0)
            {
                var drv = selectedCells[0].Item as DataRowView;
                filePath = drv?.Row["FilePath"]?.ToString();
            }
            else if (AttributesDataGrid.CurrentCell.Item != null)
            {
                var drv = AttributesDataGrid.CurrentCell.Item as DataRowView;
                filePath = drv?.Row["FilePath"]?.ToString();
            }

            if (string.IsNullOrWhiteSpace(filePath)) return;
            OpenFileInAutoCad(filePath);
        }

        private void FilesContext_OpenInAutoCAD_Click(object sender, RoutedEventArgs e)
        {
            var entry = FilesListView.SelectedItem as FileEntry;
            if (entry == null || string.IsNullOrWhiteSpace(entry.FilePath)) return;
            OpenFileInAutoCad(entry.FilePath);
        }

        private void OpenFileInAutoCad(string filePath)
        {
            try
            {
                var docManager = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager;

                // If already open, just activate it.
                foreach (Document d in docManager)
                {
                    if (string.Equals(d.Name, filePath, StringComparison.OrdinalIgnoreCase))
                    {
                        docManager.MdiActiveDocument = d;
                        BringAutoCadToFront();
                        this.WindowState = System.Windows.WindowState.Minimized;
                        return;
                    }
                }

                var activeDoc = docManager.MdiActiveDocument;
                if (activeDoc != null)
                {
                    if (!System.IO.File.Exists(filePath))
                    {
                        MessageBox.Show("Cannot find the specified file. Please verify that the file exists.",
                            "File not found", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }

                    try
                    {
                        docManager.Open(filePath, false);
                        BringAutoCadToFront();
                        this.WindowState = System.Windows.WindowState.Minimized;
                        return;
                    }
                    catch (Autodesk.AutoCAD.Runtime.Exception)
                    {
                        // Fallback to command below.
                    }

                    object oldFileDia = null;
                    object oldCmdDia = null;
                    try
                    {
                        oldFileDia = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("FILEDIA");
                        oldCmdDia = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("CMDDIA");
                        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("FILEDIA", 0);
                        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("CMDDIA", 0);

                        // Fallback: use -OPEN with dialogs disabled.
                        activeDoc.SendStringToExecute($"_.-OPEN\n\"{filePath}\"\n", true, false, false);
                    }
                    finally
                    {
                        // Delay restore to allow command to execute while dialogs are disabled.
                        _openOldFileDia = oldFileDia;
                        _openOldCmdDia = oldCmdDia;
                        ScheduleRestoreAutoCadDialogs();
                    }

                    BringAutoCadToFront();
                    this.WindowState = System.Windows.WindowState.Minimized;
                }
                else
                {
                    MessageBox.Show("No active AutoCAD document available to open the file.", "AutoCAD", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error opening file in AutoCAD: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private System.Windows.Threading.DispatcherTimer _openRestoreTimer;
        private object _openOldFileDia;
        private object _openOldCmdDia;

        private void ScheduleRestoreAutoCadDialogs()
        {
            // Restore FILEDIA/CMDDIA after the command is likely consumed.
            if (_openRestoreTimer == null)
            {
                _openRestoreTimer = new System.Windows.Threading.DispatcherTimer
                {
                    Interval = TimeSpan.FromSeconds(1)
                };
                _openRestoreTimer.Tick += (s, e) =>
                {
                    _openRestoreTimer.Stop();
                    try
                    {
                        if (_openOldFileDia != null)
                            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("FILEDIA", _openOldFileDia);
                        if (_openOldCmdDia != null)
                            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("CMDDIA", _openOldCmdDia);
                    }
                    catch
                    {
                        // Best-effort only.
                    }
                    finally
                    {
                        _openOldFileDia = null;
                        _openOldCmdDia = null;
                    }
                };
            }

            _openRestoreTimer.Stop();
            _openRestoreTimer.Start();
        }

        private void BringAutoCadToFront()
        {
            var acadWindow = Autodesk.AutoCAD.ApplicationServices.Application.MainWindow;
            if (acadWindow == null) return;

            try
            {
                // Bring AutoCAD to front without changing its size.
                if (acadWindow.WindowState == Autodesk.AutoCAD.Windows.Window.State.Minimized)
                    ShowWindow(acadWindow.Handle, 9); // SW_RESTORE only if minimized
                SetForegroundWindow(acadWindow.Handle);
            }
            catch
            {
                // Best-effort only.
            }
        }

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        private bool IsAttributeColumn(string colName)
        {
            return !(colName.Equals("File", StringComparison.OrdinalIgnoreCase)
                     || colName.Equals("LayoutOwner", StringComparison.OrdinalIgnoreCase)
                     || colName.Equals("BlockName", StringComparison.OrdinalIgnoreCase));
        }

        private void UpdateFileRowIndicesFromView()
        {
            if (_table == null || AttributesDataGrid == null) return;
            if (!_table.Columns.Contains("_FileRowIndex")) return;

            var counters = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            foreach (var item in AttributesDataGrid.Items)
            {
                if (item is DataRowView drv)
                {
                    var filePath = drv.Row["FilePath"]?.ToString();
                    if (string.IsNullOrWhiteSpace(filePath)) continue;
                    var index = counters.TryGetValue(filePath, out var existing) ? existing : 0;
                    counters[filePath] = index + 1;
                    drv.Row["_FileRowIndex"] = index;
                }
            }
        }

        private bool ShouldHandleGridClipboard()
        {
            if (AttributesDataGrid == null || !AttributesDataGrid.IsKeyboardFocusWithin)
                return false;

            // Let the active editor handle Ctrl+C/Ctrl+V
            if (System.Windows.Input.Keyboard.FocusedElement is TextBoxBase)
                return false;

            return true;
        }

        private bool CopySelectionToClipboard()
        {
            var cells = GetTargetCells();
            if (cells.Count == 0) return false;

            var cellMap = new Dictionary<(int row, int col), DataGridCellInfo>();
            var rowIndices = new HashSet<int>();
            var colIndices = new HashSet<int>();
            foreach (var cell in cells)
            {
                var rowIndex = AttributesDataGrid.Items.IndexOf(cell.Item);
                var colIndex = cell.Column?.DisplayIndex ?? -1;
                if (rowIndex < 0 || colIndex < 0) continue;
                cellMap[(rowIndex, colIndex)] = cell;
                rowIndices.Add(rowIndex);
                colIndices.Add(colIndex);
            }

            if (rowIndices.Count == 0 || colIndices.Count == 0) return false;

            int minRow = rowIndices.Min();
            int maxRow = rowIndices.Max();
            int minCol = colIndices.Min();
            int maxCol = colIndices.Max();

            var orderedColumns = AttributesDataGrid.Columns
                .OrderBy(c => c.DisplayIndex)
                .Where(c => c.DisplayIndex >= minCol && c.DisplayIndex <= maxCol)
                .ToList();

            var lines = new List<string>();
            for (int rowIndex = minRow; rowIndex <= maxRow; rowIndex++)
            {
                if (rowIndex < 0 || rowIndex >= AttributesDataGrid.Items.Count)
                    continue;

                var values = new List<string>();
                foreach (var col in orderedColumns)
                {
                    var key = (rowIndex, col.DisplayIndex);
                    if (cellMap.TryGetValue(key, out var cellInfo))
                    {
                        values.Add(GetCellText(cellInfo));
                    }
                    else
                    {
                        values.Add("");
                    }
                }
                lines.Add(string.Join("\t", values));
            }

            try
            {
                Clipboard.SetText(string.Join("\r\n", lines));
                return true;
            }
            catch
            {
                return false;
            }
        }

        private bool TryPasteClipboardToSelection()
        {
            string pasteValue = "";
            try
            {
                if (!Clipboard.ContainsText()) return false;
                pasteValue = Clipboard.GetText() ?? "";
            }
            catch
            {
                return false;
            }

            var cells = GetTargetCells();
            if (cells.Count == 0) return false;

            bool anyChange = false;
            bool baselineEnsured = false;
            foreach (var cell in cells)
            {
                var drv = cell.Item as DataRowView;
                var colName = cell.Column?.SortMemberPath;
                if (drv == null || string.IsNullOrWhiteSpace(colName)) continue;
                if (!IsAttributeColumn(colName)) continue;

                string cellText = null;
                try { cellText = drv.Row[colName]?.ToString(); } catch { }
                if (string.Equals(cellText, "N/A", StringComparison.OrdinalIgnoreCase)) continue;

                var filePath = drv.Row["FilePath"]?.ToString();
                var handle = drv.Row["Handle"]?.ToString();
                var blockName = drv.Row["BlockName"]?.ToString();
                if (string.IsNullOrWhiteSpace(filePath) || string.IsNullOrWhiteSpace(handle) || string.IsNullOrWhiteSpace(blockName))
                    continue;

                var key = new BlockAnalyzer.ChangeKey(filePath, handle, blockName, colName);
                var hadChange = _changes.TryGetValue(key, out var existingValue);
                var willChange = !hadChange || !string.Equals(existingValue ?? "", pasteValue ?? "", StringComparison.Ordinal);
                if (willChange)
                {
                    if (!baselineEnsured)
                    {
                        EnsureUndoBaseline();
                        baselineEnsured = true;
                    }
                    _changes[key] = pasteValue;
                    anyChange = true;
                }
            }

            if (anyChange)
            {
                RecordUndoState(cells.Count > 1 ? "Paste to cells" : "Paste");
                RebuildGridPreservingScroll();
                UpdateUndoRedoButtons();
            }
            return true;
        }

        private bool TryClearSelectedCells()
        {
            var cells = GetTargetCells();
            if (cells.Count == 0) return false;

            bool anyChange = false;
            bool baselineEnsured = false;
            foreach (var cell in cells)
            {
                var drv = cell.Item as DataRowView;
                var colName = cell.Column?.SortMemberPath;
                if (drv == null || string.IsNullOrWhiteSpace(colName)) continue;
                if (!IsAttributeColumn(colName)) continue;

                string cellText = null;
                try { cellText = drv.Row[colName]?.ToString(); } catch { }
                if (string.Equals(cellText, "N/A", StringComparison.OrdinalIgnoreCase)) continue;

                var filePath = drv.Row["FilePath"]?.ToString();
                var handle = drv.Row["Handle"]?.ToString();
                var blockName = drv.Row["BlockName"]?.ToString();
                if (string.IsNullOrWhiteSpace(filePath) || string.IsNullOrWhiteSpace(handle) || string.IsNullOrWhiteSpace(blockName))
                    continue;

                var key = new BlockAnalyzer.ChangeKey(filePath, handle, blockName, colName);
                var hadChange = _changes.TryGetValue(key, out var existingValue);
                var willChange = !hadChange || !string.Equals(existingValue ?? "", "", StringComparison.Ordinal);
                if (willChange)
                {
                    if (!baselineEnsured)
                    {
                        EnsureUndoBaseline();
                        baselineEnsured = true;
                    }
                    _changes[key] = "";
                    anyChange = true;
                }
            }

            if (anyChange)
            {
                RecordUndoState(cells.Count > 1 ? "Delete cells content" : "Delete cell content");
                RebuildGridPreservingScroll();
                UpdateUndoRedoButtons();
            }
            return true;
        }

        private string GetCellText(DataGridCellInfo cellInfo)
        {
            var drv = cellInfo.Item as DataRowView;
            var colName = cellInfo.Column?.SortMemberPath;
            if (drv == null || string.IsNullOrWhiteSpace(colName)) return "";

            if (IsAttributeColumn(colName))
            {
                var filePath = drv.Row["FilePath"]?.ToString();
                var handle = drv.Row["Handle"]?.ToString();
                var blockName = drv.Row["BlockName"]?.ToString();
                if (!string.IsNullOrWhiteSpace(filePath) && !string.IsNullOrWhiteSpace(handle) && !string.IsNullOrWhiteSpace(blockName))
                {
                    var key = new BlockAnalyzer.ChangeKey(filePath, handle, blockName, colName);
                    if (_changes.TryGetValue(key, out var changedValue))
                        return changedValue ?? "";
                }
            }

            return drv.Row[colName]?.ToString() ?? "";
        }

        private List<DataGridCellInfo> GetTargetCells()
        {
            if (AttributesDataGrid.SelectedCells.Count > 0)
                return AttributesDataGrid.SelectedCells.ToList();

            if (AttributesDataGrid.CurrentCell.Item != null)
                return new List<DataGridCellInfo> { AttributesDataGrid.CurrentCell };

            return new List<DataGridCellInfo>();
        }

        private DataGridCellInfo? GetSingleTargetCell()
        {
            if (AttributesDataGrid.CurrentCell.Item != null)
                return AttributesDataGrid.CurrentCell;

            if (AttributesDataGrid.SelectedCells.Count > 0)
                return AttributesDataGrid.SelectedCells[0];

            return null;
        }

        #endregion

        #region Excel Export

        private void ExportAllToExcel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var allFiles = _files.Select(f => f.FilePath).ToList();
                ExportToExcel(allFiles, null, "all blocks");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error exporting to Excel: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ExportSelectedToExcel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Export only what's visible in the table (selected files AND selected blocks)
                var selectedFiles = GetSelectedFilePaths();
                var selectedBlocks = GetSelectedBlocks();
                
                if (selectedFiles.Count == 0)
                {
                    MessageBox.Show("No files selected for export.", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }
                
                if (selectedBlocks.Count == 0)
                {
                    MessageBox.Show("No blocks selected for export.", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }
                
                ExportToExcel(selectedFiles, selectedBlocks, "selected blocks");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error exporting to Excel: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ExportToExcel_Click(object sender, RoutedEventArgs e)
        {
            var result = MessageBox.Show(
                this,
                "Export which data?\nYes = all blocks, No = selected blocks.",
                "Export to Excel",
                MessageBoxButton.YesNoCancel,
                MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                ExportAllToExcel_Click(sender, e);
            }
            else if (result == MessageBoxResult.No)
            {
                ExportSelectedToExcel_Click(sender, e);
            }
        }

        private void ExportToExcel(List<string> filePaths, HashSet<string> blockFilter, string description)
        {
            var saveDialog = new Microsoft.Win32.SaveFileDialog
            {
                Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*",
                FileName = "AttributeExport.xlsx",
                Title = $"Export {description} to Excel"
            };

            if (saveDialog.ShowDialog() != true) return;

            SetBusy(true, $"Exporting {description} to Excel...");

            // Collect all data for selected files (before try block so it's accessible later)
            var allRows = new List<BlockAnalyzer.AttributeInstanceRow>();
            foreach (var filePath in filePaths)
            {
                EnsureInstances(filePath);
                if (!_instancesByFile.TryGetValue(filePath, out var instances)) continue;
                
                foreach (var row in instances)
                {
                    // If blockFilter is null, export all blocks. Otherwise, filter by selected blocks.
                    if (blockFilter == null || blockFilter.Contains(row.BlockName))
                    {
                        allRows.Add(row);
                    }
                }
            }

            if (allRows.Count == 0)
            {
                SetBusy(false, null);
                MessageBox.Show("No data to export.", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            try
            {
                // Get all unique attribute tags
                var allTags = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var row in allRows)
                {
                    foreach (var attr in row.Attributes)
                    {
                        allTags.Add(attr.Key); // Dictionary<string, string> - Key is the tag
                    }
                }
                var sortedTags = allTags.OrderBy(t => t).ToList();

                // Create Excel workbook
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Attributes");

                    // Write header
                    int col = 1;
                    worksheet.Cell(1, col++).Value = "File";
                    worksheet.Cell(1, col++).Value = "Layout Owner";
                    worksheet.Cell(1, col++).Value = "Block Name";
                    worksheet.Cell(1, col++).Value = "Handle";
                    foreach (var tag in sortedTags)
                    {
                        worksheet.Cell(1, col++).Value = tag;
                    }

                    // Style header
                    var headerRange = worksheet.Range(1, 1, 1, col - 1);
                    headerRange.Style.Font.Bold = true;
                    headerRange.Style.Fill.BackgroundColor = XLColor.LightGray;
                    headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                    // Write data rows
                    int rowNum = 2;
                    foreach (var row in allRows)
                    {
                        col = 1;
                        worksheet.Cell(rowNum, col++).Value = System.IO.Path.GetFileName(row.FilePath);
                        worksheet.Cell(rowNum, col++).Value = row.LayoutName;
                        worksheet.Cell(rowNum, col++).Value = row.BlockName;
                        worksheet.Cell(rowNum, col++).Value = row.BlockHandle;

                        foreach (var tag in sortedTags)
                        {
                            // Check if this attribute has a change
                            var changeKey = new BlockAnalyzer.ChangeKey(row.FilePath, row.BlockHandle, row.BlockName, tag);
                            if (_changes.TryGetValue(changeKey, out var changedValue))
                            {
                                worksheet.Cell(rowNum, col).Value = changedValue;
                                // Highlight changed cells
                                worksheet.Cell(rowNum, col).Style.Fill.BackgroundColor = XLColor.LightGreen;
                            }
                            else
                            {
                                // Use original value from Dictionary
                                if (row.Attributes.TryGetValue(tag, out var originalValue))
                                {
                                    worksheet.Cell(rowNum, col).Value = originalValue;
                                }
                            }
                            col++;
                        }
                        rowNum++;
                    }

                    // Auto-fit columns
                    worksheet.Columns().AdjustToContents();

                    // Save workbook
                    workbook.SaveAs(saveDialog.FileName);
                }

                MessageBox.Show($"Successfully exported {allRows.Count} rows to {saveDialog.FileName}", "Export Complete", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            finally
            {
                SetBusy(false, null);
            }
        }

        #endregion

        private void RebuildGridPreservingScroll()
        {
            // Save scroll position
            var scrollViewer = FindVisualChild<ScrollViewer>(AttributesDataGrid);
            double? savedVerticalOffset = null;
            double? savedHorizontalOffset = null;
            if (scrollViewer != null)
            {
                savedVerticalOffset = scrollViewer.VerticalOffset;
                savedHorizontalOffset = scrollViewer.HorizontalOffset;
            }

            // Save column display order before rebuilding
            SaveColumnDisplayOrder();

            RebuildGrid();

            // Restore scroll position after rebuild
            if (scrollViewer != null && savedVerticalOffset.HasValue)
            {
                // Use BeginInvoke to restore after grid is rendered
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    var sv = FindVisualChild<ScrollViewer>(AttributesDataGrid);
                    if (sv != null)
                    {
                        sv.ScrollToVerticalOffset(savedVerticalOffset.Value);
                        if (savedHorizontalOffset.HasValue)
                            sv.ScrollToHorizontalOffset(savedHorizontalOffset.Value);
                    }
                    // Refresh styles after scroll position is restored to ensure all cells are styled
                    RefreshAllGridCellStyles();
                }), System.Windows.Threading.DispatcherPriority.Loaded);
            }
            else
            {
                // If no scroll position to restore, still refresh styles after a delay
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    RefreshAllGridCellStyles();
                }), System.Windows.Threading.DispatcherPriority.Loaded);
            }
        }
    }
}

