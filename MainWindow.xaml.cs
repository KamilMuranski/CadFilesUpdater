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

namespace CadFilesUpdater.Windows
{
    public partial class MainWindow : Window
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

        private LastEditedContext _lastEdited;
        // Remember last values used in Apply window (when no table edit was made)
        private string _lastApplyValue = "";
        private string _lastApplyBlockName = "";
        private string _lastApplyAttributeTag = "";
        private DataTable _table;
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
            var dialog = new OpenFileDialog
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
            var toRemove = FilesListView.SelectedItems.Cast<FileEntry>().ToList();
            if (toRemove.Count == 0) return;

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

            // Friendly headers
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
            // Defer until visuals exist.
            Dispatcher.BeginInvoke(new Action(() =>
            {
                for (int i = 0; i < AttributesDataGrid.Items.Count; i++)
                {
                    var row = (DataGridRow)AttributesDataGrid.ItemContainerGenerator.ContainerFromIndex(i);
                    if (row != null)
                        RefreshGridCellStylesForRow(row);
                }
            }), System.Windows.Threading.DispatcherPriority.Background);
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

                for (int colIndex = 0; colIndex < AttributesDataGrid.Columns.Count; colIndex++)
                {
                    var col = AttributesDataGrid.Columns[colIndex];
                    var colName = col.SortMemberPath ?? (col.Header?.ToString() ?? "");
                    if (colName == "File" || colName == "LayoutOwner" || colName == "BlockName") continue;

                    var cell = GetCell(row, colIndex);
                    if (cell == null) continue;

                    string cellValue = null;
                    try { cellValue = drv.Row[colName]?.ToString(); } catch { }

                    var key = new BlockAnalyzer.ChangeKey(filePath, handle, blockName, colName);
                    if (_changes.ContainsKey(key))
                    {
                        cell.Background = new SolidColorBrush(Color.FromRgb(0xBB, 0xF7, 0xD0)); // light green
                        cell.Foreground = Brushes.Black;
                        cell.FontStyle = FontStyles.Normal;
                    }
                    else if (string.Equals(cellValue, "N/A", StringComparison.OrdinalIgnoreCase))
                    {
                        cell.Background = new SolidColorBrush(Color.FromRgb(0xFE, 0xF9, 0xC3)); // light yellow
                        cell.Foreground = Brushes.Gray;
                        cell.FontStyle = FontStyles.Italic;
                    }
                    else
                    {
                        cell.ClearValue(DataGridCell.BackgroundProperty);
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
            // Horizontal scrolling creates/destroys cells because of column virtualization.
            // Refresh styles for currently realized cells so "changed" (green) stays correct while scrolling.
            if (Math.Abs(e.HorizontalChange) > 0.0 ||
                Math.Abs(e.ExtentWidthChange) > 0.0 ||
                Math.Abs(e.ViewportWidthChange) > 0.0)
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
            
            // If user edited in table, use those values; otherwise use remembered values from Apply window
            string initialValue = _lastEdited?.Value ?? _lastApplyValue ?? "";
            string initialBlockName = _lastEdited?.BlockName ?? _lastApplyBlockName ?? "";
            string initialAttributeTag = _lastEdited?.AttributeTag ?? _lastApplyAttributeTag ?? "";
            
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

            foreach (var fp in targetFiles)
            {
                EnsureInstances(fp);
                if (!_instancesByFile.TryGetValue(fp, out var rows)) continue;
                foreach (var row in rows)
                {
                    if (!row.BlockName.Equals(dlg.SelectedBlockName, StringComparison.OrdinalIgnoreCase)) continue;
                    if (!row.Attributes.ContainsKey(dlg.SelectedAttributeTag)) continue;
                    _changes[new BlockAnalyzer.ChangeKey(fp, row.BlockHandle, row.BlockName, dlg.SelectedAttributeTag)] = newValue;
                }
            }

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

            Activate();
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
            foreach (var fp in selectedFiles)
            {
                EnsureInstances(fp);
                if (!_instancesByFile.TryGetValue(fp, out var rows)) continue;
                foreach (var r in rows)
                {
                    if (!selectedBlocks.Contains(r.BlockName)) continue;
                    visibleRows.Add(r);
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
            RefreshAllGridCellStyles();
        }

        private void UpdateSummaries()
        {
            FilesSummaryText.Text = $"{_files.Count} file(s) loaded, {_files.Count(f => f.IsSelected)} selected";
        }
    }
}

