using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Media;
using Microsoft.Win32;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Windows;
using ClosedXML.Excel;

namespace CadFilesUpdater.Windows
{
    public partial class MainWindow : System.Windows.Window
    {
        private sealed class FileEntry
        {
            public string FilePath { get; set; }
            public bool IsSelected { get; set; }
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
        private int? _lastFileClickIndex;
        private int? _lastBlockClickIndex;
        private bool? _lastFileRangeState;
        private bool? _lastBlockRangeState;
        private bool _suppressSelectionEvents;
        private System.Windows.Threading.DispatcherTimer _refreshTimer;
        private bool _refreshBlocks;
        private bool _refreshGrid;
        private System.Windows.Threading.DispatcherTimer _gridStyleRefreshTimer;

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

            var existing = new HashSet<string>(_files.Select(f => f.FilePath), StringComparer.OrdinalIgnoreCase);
            var addedAny = false;
            foreach (var fp in dialog.FileNames)
            {
                if (existing.Contains(fp)) continue;
                _files.Add(new FileEntry { FilePath = fp, IsSelected = true });
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
        }

        private void RemoveFiles_Click(object sender, RoutedEventArgs e)
        {
            // Get selected files - check both SelectedItems and IsSelected property
            var toRemove = _files.Where(f => f.IsSelected).ToList();
            if (toRemove.Count == 0)
            {
                // Fallback: try SelectedItems if nothing is selected via IsSelected
                toRemove = FilesListView.SelectedItems.Cast<FileEntry>().ToList();
            }
            
            if (toRemove.Count == 0) return;

            // Record undo state before removing
            RecordUndoState("Remove files", removedFiles: toRemove.Select(f => f.FilePath).ToList());

            foreach (var fe in toRemove)
            {
                _files.Remove(fe);
                _instancesByFile.Remove(fe.FilePath);
            }

            // Drop changes for removed files
            var removedPaths = new HashSet<string>(toRemove.Select(t => t.FilePath), StringComparer.OrdinalIgnoreCase);
            var keysToRemove = _changes.Keys.Where(k => removedPaths.Contains(k.FilePath)).ToList();
            foreach (var k in keysToRemove) _changes.Remove(k);

            RebuildBlocksList();
            RebuildGrid();
            UpdateSummaries();
            UpdateUndoRedoButtons();
        }

        private void SelectAllFiles_Click(object sender, RoutedEventArgs e)
        {
            foreach (var f in _files) f.IsSelected = true;
            FilesListView.Items.Refresh();

            EnsureInstancesForSelectedFiles();
            RebuildBlocksList();
            RebuildGrid();
            UpdateSummaries();
        }

        private void UnselectAllFiles_Click(object sender, RoutedEventArgs e)
        {
            foreach (var f in _files) f.IsSelected = false;
            FilesListView.Items.Refresh();

            // No files selected => empty blocks and grid
            _blocks.Clear();
            BlocksListView.Items.Refresh();
            RebuildGrid();
            UpdateSummaries();
        }

        private void FilesSelectionChanged(object sender, RoutedEventArgs e)
        {
            if (_suppressSelectionEvents) return;

            // If this came from a checkbox click, support Shift-range selection too.
            var cb = sender as CheckBox;
            if (cb != null)
            {
                var lvi = ItemsControl.ContainerFromElement(FilesListView, cb) as ListViewItem;
                var idx = lvi != null ? FilesListView.ItemContainerGenerator.IndexFromContainer(lvi) : -1;
                var isShift = (System.Windows.Input.Keyboard.Modifiers & System.Windows.Input.ModifierKeys.Shift) == System.Windows.Input.ModifierKeys.Shift;
                var entry = lvi != null ? (lvi.DataContext as FileEntry) : null;
                var currentState = entry != null && entry.IsSelected;

                if (isShift && _lastFileClickIndex.HasValue && _lastFileRangeState.HasValue && idx >= 0)
                {
                    int start = Math.Min(_lastFileClickIndex.Value, idx);
                    int end = Math.Max(_lastFileClickIndex.Value, idx);
                    try
                    {
                        _suppressSelectionEvents = true;
                        for (int i = start; i <= end && i < _files.Count; i++)
                            _files[i].IsSelected = _lastFileRangeState.Value;
                        FilesListView.Items.Refresh();
                    }
                    finally
                    {
                        _suppressSelectionEvents = false;
                    }
                }
                if (idx >= 0)
                {
                    if (!isShift)
                        _lastFileRangeState = currentState;
                    _lastFileClickIndex = idx;
                }
            }

            // Avoid heavy work on every click; just schedule refresh.
            ScheduleRefresh(blocks: true, grid: true);
        }

        private void FilesListView_PreviewMouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (IsClickInsideCheckBox(e.OriginalSource as DependencyObject)) return;

            var lvi = FindAncestor<ListViewItem>(e.OriginalSource as DependencyObject);
            if (lvi == null) return;

            var entry = lvi.DataContext as FileEntry;
            if (entry == null) return;

            var idx = FilesListView.ItemContainerGenerator.IndexFromContainer(lvi);
            var isShift = (System.Windows.Input.Keyboard.Modifiers & System.Windows.Input.ModifierKeys.Shift) == System.Windows.Input.ModifierKeys.Shift;

            if (isShift && _lastFileClickIndex.HasValue && idx >= 0)
            {
                int start = Math.Min(_lastFileClickIndex.Value, idx);
                int end = Math.Max(_lastFileClickIndex.Value, idx);
                var desired = _lastFileRangeState ?? true;
                for (int i = start; i <= end && i < _files.Count; i++)
                    _files[i].IsSelected = desired; // shift-click = select/unselect range based on anchor
            }
            else
            {
                entry.IsSelected = !entry.IsSelected;
                _lastFileRangeState = entry.IsSelected;
            }

            if (idx >= 0) _lastFileClickIndex = idx;
            FilesListView.Items.Refresh();
            ScheduleRefresh(blocks: true, grid: true);
            e.Handled = true;
        }

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
                string.Equals(e.PropertyName, "Handle", StringComparison.OrdinalIgnoreCase))
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
                e.Column.Header = e.PropertyName; // attribute tag
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
            if (string.Equals(original ?? "", newValue ?? "", StringComparison.Ordinal))
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

                // Determine file index for alternating colors (only for non-editable columns)
                int fileIndex = _fileOrder.IndexOf(filePath);
                bool isEvenFile = fileIndex >= 0 && (fileIndex % 2 == 0);
                // Light yellow for even files, slightly darker yellow for odd files
                Color nonEditableBgColor = isEvenFile 
                    ? Color.FromRgb(0xFE, 0xF9, 0xC3) // Light yellow (original)
                    : Color.FromRgb(0xF9, 0xF3, 0xA8); // Slightly darker yellow

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
                    
                    // Set background for non-editable columns (File, LayoutOwner, BlockName) - alternating yellow
                    if (colName == "File" || colName == "LayoutOwner" || colName == "BlockName")
                    {
                        cell.Background = new SolidColorBrush(nonEditableBgColor);
                        cell.IsEnabled = false;
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
                        // N/A cell: light yellow
                        cell.Background = new SolidColorBrush(Color.FromRgb(0xFE, 0xF9, 0xC3)); // light yellow
                        cell.Foreground = Brushes.Gray;
                        cell.FontStyle = FontStyles.Italic;
                    }
                    else
                    {
                        // Normal editable cell: white background
                        cell.Background = Brushes.White;
                        cell.ClearValue(DataGridCell.ForegroundProperty);
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
                    _changes[key] = newValue;
                }
            }
            
            // Record undo state for bulk edit
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
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[MainWindow] AnalyzeAttributeInstances error: {filePath}: {ex.Message}");
                _instancesByFile[filePath] = new List<BlockAnalyzer.AttributeInstanceRow>();
            }
        }

        private List<string> GetSelectedFilePaths()
        {
            return _files.Where(f => f.IsSelected).Select(f => f.FilePath).ToList();
        }

        private HashSet<string> GetSelectedBlocks()
        {
            return _blocks.Where(b => b.IsSelected).Select(b => b.BlockName).ToHashSet(StringComparer.OrdinalIgnoreCase);
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

            var oldSelected = _blocks.Where(b => b.IsSelected).Select(b => b.BlockName).ToHashSet(StringComparer.OrdinalIgnoreCase);
            var oldUnselected = _blocks.Where(b => !b.IsSelected).Select(b => b.BlockName).ToHashSet(StringComparer.OrdinalIgnoreCase);
            _blocks.Clear();
            foreach (var bn in blockNames.OrderBy(x => x))
            {
                // Requirement: when new file(s) are selected, newly appearing blocks should be selected by default.
                // Keep user's previous choice for blocks that already existed.
                bool existedBefore = oldSelected.Contains(bn) || oldUnselected.Contains(bn);
                bool isSelected = existedBefore ? oldSelected.Contains(bn) : true;
                _blocks.Add(new BlockEntry { BlockName = bn, IsSelected = isSelected });
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

            foreach (var tag in attributeTags.OrderBy(x => x))
                _table.Columns.Add(tag, typeof(string));

            foreach (var r in visibleRows)
            {
                var dr = _table.NewRow();
                dr["File"] = System.IO.Path.GetFileNameWithoutExtension(r.FilePath);
                dr["LayoutOwner"] = r.LayoutName;
                dr["BlockName"] = r.BlockName;
                dr["FilePath"] = r.FilePath;
                dr["Handle"] = r.BlockHandle;

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
            FilesSummaryText.Text = $"{_files.Count} file(s) loaded, {_files.Count(f => f.IsSelected)} selected";
        }

        private void MainWindow_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
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

        #region Undo/Redo

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

        private void AttributesDataGrid_ContextMenuOpening(object sender, ContextMenuEventArgs e)
        {
            // Context menu is already defined in XAML, this handler can be used for dynamic updates if needed
        }

        private void ContextMenu_RevertAttribute_Click(object sender, RoutedEventArgs e)
        {
            var selectedCells = AttributesDataGrid.SelectedCells;
            if (selectedCells.Count == 0) return;

            RecordUndoState("Revert attribute(s)");

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

            foreach (var key in cellsToRevert)
                _changes.Remove(key);

            RebuildGridPreservingScroll();
            UpdateUndoRedoButtons();
        }

        private void ContextMenu_RevertBlock_Click(object sender, RoutedEventArgs e)
        {
            // Get cells from selection or current cell if nothing selected
            var cells = AttributesDataGrid.SelectedCells;
            if (cells.Count == 0 && AttributesDataGrid.CurrentCell.Item == null) return;

            RecordUndoState("Revert block(s)");

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

            foreach (var key in keysToRemove)
                _changes.Remove(key);

            RebuildGridPreservingScroll();
            UpdateUndoRedoButtons();
        }

        private void ContextMenu_RevertFile_Click(object sender, RoutedEventArgs e)
        {
            var selectedCells = AttributesDataGrid.SelectedCells;
            if (selectedCells.Count == 0) return;

            RecordUndoState("Revert file(s)");

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

            foreach (var key in keysToRemove)
                _changes.Remove(key);

            RebuildGridPreservingScroll();
            UpdateUndoRedoButtons();
        }

        private void ContextMenu_OpenInAutoCAD_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var selectedCells = AttributesDataGrid.SelectedCells;
                if (selectedCells.Count == 0 && AttributesDataGrid.CurrentCell.Item == null) return;

                var filePath = "";
                if (selectedCells.Count > 0)
                {
                    var drv = selectedCells[0].Item as DataRowView;
                    if (drv != null)
                        filePath = drv.Row["FilePath"]?.ToString();
                }
                else if (AttributesDataGrid.CurrentCell.Item != null)
                {
                    var drv = AttributesDataGrid.CurrentCell.Item as DataRowView;
                    if (drv != null)
                        filePath = drv.Row["FilePath"]?.ToString();
                }

                if (string.IsNullOrWhiteSpace(filePath)) return;

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

                // Try to open directly via DocumentManager.Open (no dialog).
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
                    catch (Autodesk.AutoCAD.Runtime.Exception ex)
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
                    LogOpenInAutoCad("No active document available to send OPEN command.");
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
                // Restore and bring AutoCAD to front.
                ShowWindow(acadWindow.Handle, 9); // SW_RESTORE
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

