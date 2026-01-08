using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using CadFilesUpdater;

namespace CadFilesUpdater.Windows
{
    public partial class MainWindow : Window
    {
        private List<string> _selectedFiles = new List<string>();
        private List<BlockInfo> _allBlocks = new List<BlockInfo>();
        private List<BlockInfo> _filteredBlocks = new List<BlockInfo>();
        private List<string> _allAttributes = new List<string>();
        private List<string> _filteredAttributes = new List<string>();
        private string _selectedBlockName = null;
        private string _selectedAttributeName = null;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void SelectFiles_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Multiselect = true,
                Filter = "AutoCAD files (*.dwg)|*.dwg|All files (*.*)|*.*",
                Title = "Select AutoCAD files"
            };

            if (dialog.ShowDialog() == true)
            {
                _selectedFiles = dialog.FileNames.ToList();
                FilesListBox.ItemsSource = _selectedFiles.Select(f => System.IO.Path.GetFileName(f));
                FilesCountText.Text = $"{_selectedFiles.Count} file(s) selected";

                // Analyze files
                try
                {
                    System.Diagnostics.Debug.WriteLine($"[MainWindow] Rozpoczynam analizę {_selectedFiles.Count} plików");
                    _allBlocks = BlockAnalyzer.AnalyzeFiles(_selectedFiles);
                    System.Diagnostics.Debug.WriteLine($"[MainWindow] Otrzymano {_allBlocks.Count} bloków z analizy");
                    _filteredBlocks = _allBlocks.ToList();
                    UpdateBlocksList();
                    System.Diagnostics.Debug.WriteLine($"[MainWindow] Zaktualizowano listę bloków. Liczba w liście: {BlocksListBox.Items.Count}");
                    
                    if (_allBlocks.Count == 0)
                    {
                        MessageBox.Show("No dynamic blocks found in selected files.", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[MainWindow] Error during analysis: {ex.Message}");
                    System.Diagnostics.Debug.WriteLine($"[MainWindow] StackTrace: {ex.StackTrace}");
                    MessageBox.Show($"Error analyzing files: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void BlockSearchBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            var searchText = BlockSearchBox.Text.ToLower();
            System.Diagnostics.Debug.WriteLine($"[MainWindow] Wyszukiwanie bloków: '{searchText}' (wszystkich bloków: {_allBlocks.Count})");
            
            if (string.IsNullOrWhiteSpace(searchText))
            {
                _filteredBlocks = _allBlocks.ToList();
            }
            else
            {
                _filteredBlocks = _allBlocks
                    .Where(b => b.BlockName.ToLower().Contains(searchText))
                    .ToList();
            }
            
            System.Diagnostics.Debug.WriteLine($"[MainWindow] Znaleziono {_filteredBlocks.Count} bloków pasujących do '{searchText}'");
            UpdateBlocksList();
        }

        private void UpdateBlocksList()
        {
            var blockNames = _filteredBlocks.Select(b => b.BlockName).OrderBy(n => n).ToList();
            System.Diagnostics.Debug.WriteLine($"[MainWindow] Aktualizuję listę bloków. Liczba: {blockNames.Count}");
            BlocksListBox.ItemsSource = blockNames;
        }

        private void BlocksListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (BlocksListBox.SelectedItem != null)
            {
                _selectedBlockName = BlocksListBox.SelectedItem.ToString();
                var selectedBlock = _allBlocks.FirstOrDefault(b => b.BlockName == _selectedBlockName);
                
                if (selectedBlock != null)
                {
                    _allAttributes = selectedBlock.Attributes.OrderBy(a => a).ToList();
                    _filteredAttributes = _allAttributes.ToList();
                    UpdateAttributesList();
                    AttributeSearchBox.IsEnabled = true;
                    AttributesListBox.IsEnabled = true;
                }
            }
            else
            {
                _selectedBlockName = null;
                _allAttributes.Clear();
                _filteredAttributes.Clear();
                UpdateAttributesList();
                AttributeSearchBox.IsEnabled = false;
                AttributesListBox.IsEnabled = false;
                ValueTextBox.IsEnabled = false;
                SubmitButton.IsEnabled = false;
            }
        }

        private void AttributeSearchBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            var searchText = AttributeSearchBox.Text.ToLower();
            if (string.IsNullOrWhiteSpace(searchText))
            {
                _filteredAttributes = _allAttributes.ToList();
            }
            else
            {
                _filteredAttributes = _allAttributes
                    .Where(a => a.ToLower().Contains(searchText))
                    .ToList();
            }
            UpdateAttributesList();
        }

        private void UpdateAttributesList()
        {
            AttributesListBox.ItemsSource = _filteredAttributes;
        }

        private void AttributesListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (AttributesListBox.SelectedItem != null)
            {
                _selectedAttributeName = AttributesListBox.SelectedItem.ToString();
                ValueTextBox.IsEnabled = true;
                SubmitButton.IsEnabled = !string.IsNullOrWhiteSpace(ValueTextBox.Text);
            }
            else
            {
                _selectedAttributeName = null;
                ValueTextBox.IsEnabled = false;
                SubmitButton.IsEnabled = false;
            }
        }

        private void ValueTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            SubmitButton.IsEnabled = !string.IsNullOrWhiteSpace(ValueTextBox.Text) && 
                                     !string.IsNullOrEmpty(_selectedAttributeName);
        }

        private void SubmitButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(_selectedBlockName) || 
                string.IsNullOrWhiteSpace(_selectedAttributeName) ||
                string.IsNullOrWhiteSpace(ValueTextBox.Text))
            {
                MessageBox.Show("Please select a block, attribute and enter a value.", "Missing data", 
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var result = MessageBox.Show(
                $"Are you sure you want to update all blocks '{_selectedBlockName}' " +
                $"with attribute '{_selectedAttributeName}' to value '{ValueTextBox.Text}' " +
                $"in {_selectedFiles.Count} file(s)?",
                "Confirm operation",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                try
                {
                    this.Cursor = System.Windows.Input.Cursors.Wait;
                    SubmitButton.IsEnabled = false;

                    bool success = BlockAnalyzer.UpdateBlocksInFiles(
                        _selectedFiles,
                        _selectedBlockName,
                        _selectedAttributeName,
                        ValueTextBox.Text);

                    if (success)
                    {
                        MessageBox.Show(
                            $"Operation completed successfully!\nUpdated {_selectedFiles.Count} file(s).",
                            "Success",
                            MessageBoxButton.OK,
                            MessageBoxImage.Information);
                    }
                    else
                    {
                        MessageBox.Show(
                            "Operation completed with errors. Check console for details.",
                            "Warning",
                            MessageBoxButton.OK,
                            MessageBoxImage.Warning);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(
                        $"Error during update: {ex.Message}",
                        "Error",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);
                }
                finally
                {
                    this.Cursor = System.Windows.Input.Cursors.Arrow;
                    SubmitButton.IsEnabled = true;
                }
            }
        }
    }
}
