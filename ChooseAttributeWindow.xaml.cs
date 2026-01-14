using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Media;

namespace CadFilesUpdater.Windows
{
    public partial class ChooseAttributeWindow : Window
    {
        private sealed class FileEntry
        {
            public string FilePath { get; set; }
            public bool IsSelected { get; set; }
        }

        private readonly ObservableCollection<FileEntry> _files = new ObservableCollection<FileEntry>();
        private readonly List<string> _allFilePaths;
        private readonly Func<string, List<BlockAnalyzer.AttributeInstanceRow>> _instanceLoader;
        private readonly Dictionary<string, List<BlockAnalyzer.AttributeInstanceRow>> _cache =
            new Dictionary<string, List<BlockAnalyzer.AttributeInstanceRow>>(StringComparer.OrdinalIgnoreCase);

        public IReadOnlyList<string> SelectedFilePaths
        {
            get
            {
                var selected = FilesList?.SelectedItem as FileEntry;
                return selected != null ? new List<string> { selected.FilePath } : new List<string>();
            }
        }
        public string SelectedBlockName { get; private set; }
        public string SelectedAttributeTag { get; private set; }

        public ChooseAttributeWindow(
            List<string> allFilePaths,
            List<string> defaultSelectedFilePaths,
            string initialBlockName,
            string initialAttributeTag,
            Func<string, List<BlockAnalyzer.AttributeInstanceRow>> instanceLoader)
        {
            InitializeComponent();

            _allFilePaths = allFilePaths ?? new List<string>();
            _instanceLoader = instanceLoader;

            // Don't select any files by default - show all blocks
            foreach (var fp in _allFilePaths)
                _files.Add(new FileEntry { FilePath = fp, IsSelected = false });

            FilesList.ItemsSource = _files;

            SelectedBlockName = initialBlockName;
            SelectedAttributeTag = initialAttributeTag;

            RebuildBlocks();
            if (!string.IsNullOrWhiteSpace(SelectedBlockName))
                BlocksList.SelectedItem = SelectedBlockName;

            RebuildAttributes();
            if (!string.IsNullOrWhiteSpace(SelectedAttributeTag))
                AttributesList.SelectedItem = SelectedAttributeTag;
        }

        private void AllFiles_Click(object sender, RoutedEventArgs e)
        {
            // Unselect all files and show blocks from all files
            FilesList.SelectedItem = null;
            RebuildBlocks();
            RebuildAttributes();
        }

        private void FilesList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            RebuildBlocks();
            RebuildAttributes();
        }

        private void BlocksList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            SelectedBlockName = BlocksList.SelectedItem as string;
            RebuildAttributes();
        }

        private void Ok_Click(object sender, RoutedEventArgs e)
        {
            SelectedBlockName = BlocksList.SelectedItem as string;
            SelectedAttributeTag = AttributesList.SelectedItem as string;

            // No need to require file selection - we can work with all files
            if (string.IsNullOrWhiteSpace(SelectedBlockName))
            {
                MessageBox.Show(this, "Select a block.", "Validation", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            if (string.IsNullOrWhiteSpace(SelectedAttributeTag))
            {
                MessageBox.Show(this, "Select an attribute.", "Validation", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            DialogResult = true;
            Close();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void RebuildBlocks()
        {
            var blocks = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var selectedPaths = SelectedFilePaths;
            
            // If no files selected, show blocks from all files
            if (selectedPaths.Count == 0)
            {
                foreach (var fp in _allFilePaths)
                {
                    foreach (var r in LoadInstances(fp))
                        blocks.Add(r.BlockName);
                }
            }
            else
            {
                // Show blocks only from selected file
                foreach (var fp in selectedPaths)
                {
                    foreach (var r in LoadInstances(fp))
                        blocks.Add(r.BlockName);
                }
            }
            
            BlocksList.ItemsSource = blocks.OrderBy(x => x).ToList();
            // Clear block selection when list changes
            BlocksList.SelectedItem = null;
        }

        private void RebuildAttributes()
        {
            var blockName = BlocksList.SelectedItem as string;
            var attrs = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (!string.IsNullOrWhiteSpace(blockName))
            {
                var selectedPaths = SelectedFilePaths;
                // If no files selected, check all files
                var filesToCheck = selectedPaths.Count == 0 ? _allFilePaths : selectedPaths;
                
                foreach (var fp in filesToCheck)
                {
                    foreach (var r in LoadInstances(fp))
                    {
                        if (!r.BlockName.Equals(blockName, StringComparison.OrdinalIgnoreCase)) continue;
                        foreach (var tag in r.Attributes.Keys)
                            attrs.Add(tag);
                    }
                }
            }
            AttributesList.ItemsSource = attrs.OrderBy(x => x).ToList();
        }

        private List<BlockAnalyzer.AttributeInstanceRow> LoadInstances(string filePath)
        {
            if (_cache.TryGetValue(filePath, out var rows)) return rows;
            try
            {
                rows = _instanceLoader != null ? (_instanceLoader(filePath) ?? new List<BlockAnalyzer.AttributeInstanceRow>()) : new List<BlockAnalyzer.AttributeInstanceRow>();
            }
            catch
            {
                rows = new List<BlockAnalyzer.AttributeInstanceRow>();
            }
            _cache[filePath] = rows;
            return rows;
        }
    }
}

