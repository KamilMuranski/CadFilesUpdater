using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using Microsoft.Win32;
using CadFilesUpdater;
using CadFilesUpdater.Windows;

namespace CadFilesUpdater.Windows
{
    public partial class MainWindow : Window
    {
        private List<string> _selectedFiles = new List<string>();
        private ObservableCollection<string> _filePathsCollection = new ObservableCollection<string>();
        private List<BlockInfo> _allBlocks = new List<BlockInfo>();
        private List<BlockInfo> _filteredBlocks = new List<BlockInfo>();
        private List<string> _allAttributes = new List<string>();
        private List<string> _filteredAttributes = new List<string>();
        private string _selectedBlockName = null;
        private string _selectedAttributeName = null;
        private List<string> _selectedFilesForFiltering = new List<string>(); // Currently selected files for filtering (empty = all files)

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
                System.Diagnostics.Debug.WriteLine($"[MainWindow] Załadowano {_selectedFiles.Count} plików");
                System.Diagnostics.Debug.WriteLine($"[MainWindow] Lista plików: {string.Join(", ", _selectedFiles.Select(f => System.IO.Path.GetFileName(f)))}");
                
                // Clear existing ItemsSource first
                FilesListBox.ItemsSource = null;
                FilesListBox.Items.Clear();
                
                // Clear and populate ObservableCollection with full paths (to avoid duplicates)
                _filePathsCollection.Clear();
                foreach (var filePath in _selectedFiles)
                {
                    _filePathsCollection.Add(filePath);
                }
                
                // Set ItemsSource to ObservableCollection with full paths
                FilesListBox.ItemsSource = _filePathsCollection;
                FilesCountText.Text = $"{_selectedFiles.Count} file(s) selected";
                _selectedFilesForFiltering.Clear(); // Reset selection - no files selected initially
                
                // Force refresh of ListBox
                FilesListBox.UpdateLayout();
                FilesListBox.InvalidateVisual();
                
                // Wait for UI to update and then check
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    System.Diagnostics.Debug.WriteLine($"[MainWindow] Ustawiono ItemsSource. Liczba elementów w FilesListBox: {FilesListBox.Items.Count}");
                    System.Diagnostics.Debug.WriteLine($"[MainWindow] Liczba elementów w ObservableCollection: {_filePathsCollection.Count}");
                    
                    // Verify all items are visible
                    if (FilesListBox.Items.Count != _selectedFiles.Count)
                    {
                        System.Diagnostics.Debug.WriteLine($"[MainWindow] BŁĄD: Liczba elementów w ListBox ({FilesListBox.Items.Count}) nie zgadza się z liczbą plików ({_selectedFiles.Count})!");
                    }
                }), System.Windows.Threading.DispatcherPriority.Loaded);

                // Analyze files
                try
                {
                    System.Diagnostics.Debug.WriteLine($"[MainWindow] Rozpoczynam analizę {_selectedFiles.Count} plików");
                    _allBlocks = BlockAnalyzer.AnalyzeFiles(_selectedFiles);
                    System.Diagnostics.Debug.WriteLine($"[MainWindow] Otrzymano {_allBlocks.Count} bloków z analizy");
                    
                    // On initial load, select all files to show all blocks
                    FilesListBox.SelectAll();
                    // This will trigger FilesListBox_SelectionChanged which will call ApplyFilters and UpdateBlocksList
                    System.Diagnostics.Debug.WriteLine($"[MainWindow] Zaznaczono wszystkie pliki na początku");
                    
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

        private void SelectAllFiles_Click(object sender, RoutedEventArgs e)
        {
            // Select all files in the list
            FilesListBox.SelectAll();
        }

        private void UnselectAllFiles_Click(object sender, RoutedEventArgs e)
        {
            // Clear file selection - when no files are selected, show no blocks
            FilesListBox.SelectedItems.Clear();
            _selectedFilesForFiltering.Clear();
            ApplyFilters();
            UpdateBlocksList();
        }

        private void FilesListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Get all selected file paths (ItemsSource contains full paths now)
            var selectedFilePaths = FilesListBox.SelectedItems.Cast<string>().ToList();
            
            // Use selected file paths directly
            _selectedFilesForFiltering = selectedFilePaths;
            
            var selectedFileNames = selectedFilePaths.Select(f => System.IO.Path.GetFileName(f)).ToList();
            System.Diagnostics.Debug.WriteLine($"[MainWindow] Zaznaczono {_selectedFilesForFiltering.Count} plików: {string.Join(", ", selectedFileNames)}");
            
            // Clear block selection and attributes when file selection changes
            // This must be done before updating blocks list
            BlocksListBox.SelectedItem = null;
            _selectedBlockName = null;
            _selectedAttributeName = null;
            _allAttributes.Clear();
            _filteredAttributes.Clear();
            AttributesListBox.ItemsSource = null; // Clear the view immediately
            AttributeSearchBox.Text = ""; // Clear search text
            AttributeSearchBox.IsEnabled = false;
            AttributesListBox.IsEnabled = false;
            AttributeSearchPlaceholder.Visibility = Visibility.Visible;
            ValueTextBox.Text = "";
            ValueTextBox.IsEnabled = false;
            
            // Re-apply filters with new file selection
            ApplyFilters();
            UpdateBlocksList();
            // UpdateBlocksList will handle automatic selection if there's only one block
        }

        private void BlockSearchBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            // Show/hide placeholder based on text content
            BlockSearchPlaceholder.Visibility = string.IsNullOrEmpty(BlockSearchBox.Text) ? Visibility.Visible : Visibility.Collapsed;
            
            // Re-apply filters (includes search text)
            ApplyFilters();
            UpdateBlocksList();
        }

        private void ApplyFilters()
        {
            // Start with all blocks
            var filtered = _allBlocks.ToList();
            
            // Filter: only blocks with at least one attribute
            filtered = filtered.Where(b => b.Attributes != null && b.Attributes.Count > 0).ToList();
            System.Diagnostics.Debug.WriteLine($"[MainWindow] Po filtrowaniu bloków z atrybutami: {filtered.Count} bloków");
            
            // Filter: by selected files (if files are selected)
            // If no files are selected, show no blocks (empty list)
            if (_selectedFilesForFiltering != null && _selectedFilesForFiltering.Count > 0)
            {
                filtered = filtered.Where(b => 
                    b.FilePath != null && 
                    _selectedFilesForFiltering.Any(selectedFile => 
                        b.FilePath.Equals(selectedFile, StringComparison.OrdinalIgnoreCase)))
                    .ToList();
                System.Diagnostics.Debug.WriteLine($"[MainWindow] Po filtrowaniu po {_selectedFilesForFiltering.Count} plikach: {filtered.Count} bloków");
            }
            else
            {
                // No files selected - show no blocks
                filtered.Clear();
                System.Diagnostics.Debug.WriteLine($"[MainWindow] Brak zaznaczonych plików - lista bloków pusta");
            }
            
            // Filter: by search text (if search box has text)
            var searchText = BlockSearchBox.Text?.ToLower();
            if (!string.IsNullOrWhiteSpace(searchText))
            {
                filtered = filtered.Where(b => b.BlockName.ToLower().Contains(searchText)).ToList();
                System.Diagnostics.Debug.WriteLine($"[MainWindow] Po filtrowaniu po tekście '{searchText}': {filtered.Count} bloków");
            }
            
            _filteredBlocks = filtered;
            
            // Clear selection when filters change - this also clears attributes
            BlocksListBox.SelectedItem = null;
            _selectedBlockName = null;
            _selectedAttributeName = null;
            _allAttributes.Clear();
            _filteredAttributes.Clear();
            UpdateAttributesList();
            AttributeSearchBox.IsEnabled = false;
            AttributesListBox.IsEnabled = false;
            ValueTextBox.Text = "";
            ValueTextBox.IsEnabled = false;
        }

        private void BlockSearchBox_GotFocus(object sender, RoutedEventArgs e)
        {
            BlockSearchPlaceholder.Visibility = Visibility.Collapsed;
        }

        private void BlockSearchBox_LostFocus(object sender, RoutedEventArgs e)
        {
            BlockSearchPlaceholder.Visibility = string.IsNullOrEmpty(BlockSearchBox.Text) ? Visibility.Visible : Visibility.Collapsed;
        }

        private void UpdateBlocksList()
        {
            // Get unique block names (blocks can appear in multiple files, but we want unique names)
            var blockNames = _filteredBlocks
                .Select(b => b.BlockName)
                .Distinct()
                .OrderBy(n => n)
                .ToList();
            
            System.Diagnostics.Debug.WriteLine($"[MainWindow] Aktualizuję listę bloków. Liczba unikalnych nazw: {blockNames.Count} (z {_filteredBlocks.Count} bloków)");
            
            // Temporarily unsubscribe to prevent SelectionChanged during update
            BlocksListBox.SelectionChanged -= BlocksListBox_SelectionChanged;
            
            try
            {
                // Completely clear selection and items source to prevent index mismatch
                BlocksListBox.SelectedIndex = -1;
                BlocksListBox.SelectedItem = null;
                BlocksListBox.ItemsSource = null;
                BlocksListBox.UpdateLayout();
                
                // Set new ItemsSource
                BlocksListBox.ItemsSource = blockNames;
                System.Diagnostics.Debug.WriteLine($"[MainWindow] Ustawiono ItemsSource. Liczba elementów w BlocksListBox: {BlocksListBox.Items.Count}");
            }
            finally
            {
                // Re-subscribe to SelectionChanged FIRST
                BlocksListBox.SelectionChanged += BlocksListBox_SelectionChanged;
                
                // If there's only one block, automatically select it AFTER re-subscribing
                // This will trigger SelectionChanged which will load attributes
                if (blockNames.Count == 1)
                {
                    BlocksListBox.SelectedItem = blockNames[0];
                    System.Diagnostics.Debug.WriteLine($"[MainWindow] Automatycznie zaznaczono jedyny dostępny blok: {blockNames[0]}");
                    // SelectionChanged will be triggered automatically and will load attributes
                }
                else
                {
                    // Make sure attributes are cleared if no blocks or multiple blocks
                    _allAttributes.Clear();
                    _filteredAttributes.Clear();
                    AttributesListBox.ItemsSource = null;
                    AttributeSearchBox.IsEnabled = false;
                    AttributesListBox.IsEnabled = false;
                    AttributeSearchPlaceholder.Visibility = Visibility.Visible;
                    ValueTextBox.Text = "";
                    ValueTextBox.IsEnabled = false;
                }
            }
        }

        private void BlocksListBox_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            var listBox = sender as ListBox;
            if (listBox != null)
            {
                // Get the mouse position relative to the ListBox
                var point = e.GetPosition(listBox);
                
                // Use HitTest to find the element at mouse position
                var hitResult = VisualTreeHelper.HitTest(listBox, point);
                if (hitResult != null && hitResult.VisualHit != null)
                {
                    // Walk up the visual tree to find the ListBoxItem container
                    DependencyObject container = hitResult.VisualHit;
                    ListBoxItem clickedItem = null;
                    
                    while (container != null && container != listBox)
                    {
                        if (container is ListBoxItem item)
                        {
                            clickedItem = item;
                            break;
                        }
                        container = VisualTreeHelper.GetParent(container);
                    }
                    
                    if (clickedItem != null)
                    {
                        var clickedIndex = listBox.ItemContainerGenerator.IndexFromContainer(clickedItem);
                        System.Diagnostics.Debug.WriteLine($"[MainWindow] PreviewMouseLeftButtonDown: Kliknięto indeks {clickedIndex}, punkt=({point.X:F1}, {point.Y:F1})");
                        
                        if (clickedIndex >= 0 && clickedIndex < listBox.Items.Count)
                        {
                            // Temporarily unsubscribe to prevent SelectionChanged from firing with wrong index
                            listBox.SelectionChanged -= BlocksListBox_SelectionChanged;
                            
                            try
                            {
                                // Directly set the selection to the clicked item BEFORE default behavior
                                listBox.SelectedIndex = clickedIndex;
                                var selectedItem = listBox.SelectedItem;
                                System.Diagnostics.Debug.WriteLine($"[MainWindow] PreviewMouseLeftButtonDown: Ustawiono SelectedIndex={clickedIndex}, SelectedItem={selectedItem}");
                                
                                // Manually trigger the selection changed handler
                                if (selectedItem != null)
                                {
                                    BlocksListBox_SelectionChanged(listBox, null);
                                }
                            }
                            finally
                            {
                                // Re-subscribe to SelectionChanged after a delay to let selection settle
                                Dispatcher.BeginInvoke(new Action(() =>
                                {
                                    listBox.SelectionChanged += BlocksListBox_SelectionChanged;
                                }), System.Windows.Threading.DispatcherPriority.Input);
                            }
                            
                            // Mark event as handled to prevent default selection behavior
                            e.Handled = true;
                            return;
                        }
                    }
                }
            }
        }

        private void BlocksListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // This can be called with null e when called manually from PreviewMouseLeftButtonDown
            System.Diagnostics.Debug.WriteLine($"[MainWindow] SelectionChanged: SelectedIndex={BlocksListBox.SelectedIndex}, SelectedItem={BlocksListBox.SelectedItem}");
            
            if (BlocksListBox.SelectedItem != null)
            {
                _selectedBlockName = BlocksListBox.SelectedItem.ToString();
                System.Diagnostics.Debug.WriteLine($"[MainWindow] Wybrano blok: '{_selectedBlockName}' (indeks: {BlocksListBox.SelectedIndex})");
                
                // Find block - if files are selected, prefer blocks from those files, otherwise take first match
                BlockInfo selectedBlock = null;
                if (_selectedFilesForFiltering != null && _selectedFilesForFiltering.Count > 0)
                {
                    // Find block from selected files first
                    selectedBlock = _allBlocks.FirstOrDefault(b => 
                        b.BlockName == _selectedBlockName && 
                        b.FilePath != null && 
                        _selectedFilesForFiltering.Any(selectedFile => 
                            b.FilePath.Equals(selectedFile, StringComparison.OrdinalIgnoreCase)));
                }
                
                // If not found or no files selected, take first match from all blocks
                if (selectedBlock == null)
                {
                    selectedBlock = _allBlocks.FirstOrDefault(b => b.BlockName == _selectedBlockName);
                }
                
                if (selectedBlock != null && selectedBlock.Attributes != null && selectedBlock.Attributes.Count > 0)
                {
                    _allAttributes = selectedBlock.Attributes.OrderBy(a => a).ToList();
                    _filteredAttributes = _allAttributes.ToList();
                    UpdateAttributesList();
                    AttributeSearchBox.IsEnabled = true;
                    AttributesListBox.IsEnabled = true;
                    // Show placeholder when field is enabled and empty
                    AttributeSearchPlaceholder.Visibility = string.IsNullOrEmpty(AttributeSearchBox.Text) ? Visibility.Visible : Visibility.Collapsed;
                    // ValueTextBox remains disabled until attribute is selected
                }
                else
                {
                    // No block selected or block has no attributes - clear everything
                    System.Diagnostics.Debug.WriteLine($"[MainWindow] Brak zaznaczonego bloku lub blok bez atrybutów");
                    _selectedBlockName = null;
                    _selectedAttributeName = null;
                    _allAttributes.Clear();
                    _filteredAttributes.Clear();
                    UpdateAttributesList();
                    AttributeSearchBox.IsEnabled = false;
                    AttributesListBox.IsEnabled = false;
                    AttributeSearchPlaceholder.Visibility = Visibility.Visible;
                    ValueTextBox.Text = "";
                    ValueTextBox.IsEnabled = false;
                }
            }
            else
            {
                System.Diagnostics.Debug.WriteLine($"[MainWindow] Brak zaznaczenia");
                _selectedBlockName = null;
                _allAttributes.Clear();
                _filteredAttributes.Clear();
                UpdateAttributesList();
                AttributeSearchBox.IsEnabled = false;
                AttributesListBox.IsEnabled = false;
                ValueTextBox.IsEnabled = false;
                SubmitButton.IsEnabled = false;
                // Hide placeholder when field is disabled
                AttributeSearchPlaceholder.Visibility = Visibility.Collapsed;
            }
        }

        private void AttributeSearchBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            // Show/hide placeholder based on text content
            AttributeSearchPlaceholder.Visibility = string.IsNullOrEmpty(AttributeSearchBox.Text) ? Visibility.Visible : Visibility.Collapsed;
            
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

        private void AttributeSearchBox_GotFocus(object sender, RoutedEventArgs e)
        {
            AttributeSearchPlaceholder.Visibility = Visibility.Collapsed;
        }

        private void AttributeSearchBox_LostFocus(object sender, RoutedEventArgs e)
        {
            AttributeSearchPlaceholder.Visibility = string.IsNullOrEmpty(AttributeSearchBox.Text) ? Visibility.Visible : Visibility.Collapsed;
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
                // IMPORTANT: Copy all UI values to local variables BEFORE Task.Run
                // This ensures we're not accessing UI objects from background thread
                var selectedFiles = _selectedFiles.ToList(); // Create copy
                var selectedBlockName = _selectedBlockName;
                var selectedAttributeName = _selectedAttributeName;
                var valueText = ValueTextBox.Text; // Copy text BEFORE Task.Run
                
                System.Diagnostics.Debug.WriteLine($"[MainWindow] Copied values on UI thread: {selectedFiles.Count} files, block: {selectedBlockName}, attribute: {selectedAttributeName}, value: {valueText}");
                
                // Show progress window
                var progressWindow = new ProgressWindow
                {
                    Owner = this
                };
                progressWindow.Show();

                SubmitButton.IsEnabled = false;

                // Store reference to progress window for use in background thread
                var progressWindowRef = progressWindow;
                
                // Run update in background task
                System.Diagnostics.Debug.WriteLine($"[MainWindow] Starting Task.Run on thread: {System.Threading.Thread.CurrentThread.ManagedThreadId}");
                Task.Run(() =>
                {
                    try
                    {
                        System.Diagnostics.Debug.WriteLine($"[MainWindow] Task.Run started on thread: {System.Threading.Thread.CurrentThread.ManagedThreadId}");
                        System.Diagnostics.Debug.WriteLine($"[MainWindow] progressWindowRef is null: {progressWindowRef == null}");
                        System.Diagnostics.Debug.WriteLine($"[MainWindow] progressWindowRef.Dispatcher is null: {progressWindowRef?.Dispatcher == null}");
                        
                        System.Diagnostics.Debug.WriteLine($"[MainWindow] Copied values: {selectedFiles.Count} files, block: {selectedBlockName}, attribute: {selectedAttributeName}");
                        
                        var updateResult = BlockAnalyzer.UpdateBlocksInFiles(
                            selectedFiles,
                            selectedBlockName,
                            selectedAttributeName,
                            valueText,
                            (processed, total, currentFile) =>
                            {
                                System.Diagnostics.Debug.WriteLine($"[MainWindow] Progress callback called on thread: {System.Threading.Thread.CurrentThread.ManagedThreadId}, processed: {processed}/{total}");
                                
                                // Update progress window - must use the window's own Dispatcher
                                // Store values in local variables for thread safety
                                var proc = processed;
                                var tot = total;
                                var file = currentFile ?? "";
                                
                                System.Diagnostics.Debug.WriteLine($"[MainWindow] Progress callback: proc={proc}, tot={tot}, file={file}");
                                System.Diagnostics.Debug.WriteLine($"[MainWindow] progressWindowRef is null: {progressWindowRef == null}");
                                
                                // DO NOT access IsLoaded or any UI properties from background thread!
                                // Just try to invoke through Dispatcher - let the callback check IsLoaded
                                if (progressWindowRef != null && progressWindowRef.Dispatcher != null)
                                {
                                    try
                                    {
                                        System.Diagnostics.Debug.WriteLine($"[MainWindow] Calling Dispatcher.BeginInvoke from thread: {System.Threading.Thread.CurrentThread.ManagedThreadId}");
                                        progressWindowRef.Dispatcher.BeginInvoke(new Action(() =>
                                        {
                                            try
                                            {
                                                System.Diagnostics.Debug.WriteLine($"[MainWindow] Inside Dispatcher callback on thread: {System.Threading.Thread.CurrentThread.ManagedThreadId}");
                                                // Check IsLoaded INSIDE Dispatcher callback (on UI thread)
                                                if (progressWindowRef != null && progressWindowRef.IsLoaded)
                                                {
                                                    System.Diagnostics.Debug.WriteLine($"[MainWindow] Calling UpdateProgress: {proc}/{tot}, file: {file}");
                                                    progressWindowRef.UpdateProgress(proc, tot, file);
                                                    System.Diagnostics.Debug.WriteLine($"[MainWindow] UpdateProgress completed");
                                                }
                                                else
                                                {
                                                    System.Diagnostics.Debug.WriteLine($"[MainWindow] Progress window not loaded, skipping UpdateProgress");
                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                // Log error but don't crash
                                                System.Diagnostics.Debug.WriteLine($"[MainWindow] Error in Dispatcher callback: {ex.Message}");
                                                System.Diagnostics.Debug.WriteLine($"[MainWindow] StackTrace: {ex.StackTrace}");
                                            }
                                        }), System.Windows.Threading.DispatcherPriority.Normal);
                                        System.Diagnostics.Debug.WriteLine($"[MainWindow] Dispatcher.BeginInvoke called successfully");
                                    }
                                    catch (Exception dispatcherEx)
                                    {
                                        System.Diagnostics.Debug.WriteLine($"[MainWindow] Error invoking Dispatcher: {dispatcherEx.Message}");
                                        System.Diagnostics.Debug.WriteLine($"[MainWindow] Dispatcher Exception StackTrace: {dispatcherEx.StackTrace}");
                                    }
                                }
                                else
                                {
                                    System.Diagnostics.Debug.WriteLine($"[MainWindow] progressWindowRef or Dispatcher is null, cannot update progress");
                                }
                            });
                        
                        System.Diagnostics.Debug.WriteLine($"[MainWindow] UpdateBlocksInFiles completed. Result: {updateResult.SuccessfulFiles}/{updateResult.TotalFiles} successful, {updateResult.FailedFiles} failed");

                        // IMPORTANT: Copy all data from updateResult to local variables BEFORE Dispatcher.BeginInvoke
                        // This ensures thread safety - we're copying primitive types and strings, not UI objects
                        System.Diagnostics.Debug.WriteLine($"[MainWindow] Copying updateResult data to local variables");
                        var totalFiles = updateResult.TotalFiles;
                        var successfulFiles = updateResult.SuccessfulFiles;
                        var failedFiles = updateResult.FailedFiles;
                        var errors = updateResult.Errors.Select(error => new { 
                            FilePath = error.FilePath ?? "", 
                            ErrorMessage = error.ErrorMessage ?? "" 
                        }).ToList(); // Create anonymous objects with copied strings
                        
                        System.Diagnostics.Debug.WriteLine($"[MainWindow] Copied data: total={totalFiles}, success={successfulFiles}, failed={failedFiles}, errors={errors.Count}");
                        System.Diagnostics.Debug.WriteLine($"[MainWindow] progressWindowRef is null: {progressWindowRef == null}");
                        System.Diagnostics.Debug.WriteLine($"[MainWindow] progressWindowRef.Dispatcher is null: {progressWindowRef?.Dispatcher == null}");
                        
                        // Close progress window and show results on UI thread
                        // Use the window's Dispatcher for thread safety
                        if (progressWindowRef != null && progressWindowRef.Dispatcher != null)
                        {
                            System.Diagnostics.Debug.WriteLine($"[MainWindow] Calling Dispatcher.BeginInvoke for results on thread: {System.Threading.Thread.CurrentThread.ManagedThreadId}");
                            progressWindowRef.Dispatcher.BeginInvoke(new Action(() =>
                            {
                                System.Diagnostics.Debug.WriteLine($"[MainWindow] Inside results Dispatcher callback on thread: {System.Threading.Thread.CurrentThread.ManagedThreadId}");
                                {
                                    try
                                    {
                                        if (progressWindowRef != null)
                                        {
                                            progressWindowRef.Close();
                                        }
                                        SubmitButton.IsEnabled = true;

                                        string message = $"Operation completed!\n\n";
                                        message += $"Successfully processed: {successfulFiles}/{totalFiles} file(s)\n";

                                        if (failedFiles > 0)
                                        {
                                            message += $"\nFailed: {failedFiles} file(s)\n\n";
                                            message += "Files that could not be processed:\n";

                                            // Simply list file names - don't try to access error details that might cause thread issues
                                            foreach (var error in errors)
                                            {
                                                try
                                                {
                                                    var fileName = System.IO.Path.GetFileName(error.FilePath);
                                                    message += $"  • {fileName}\n";
                                                }
                                                catch
                                                {
                                                    // If we can't get filename, just skip this entry
                                                }
                                            }

                                            MessageBox.Show(
                                                message,
                                                successfulFiles > 0 ? "Completed with errors" : "Error",
                                                MessageBoxButton.OK,
                                                successfulFiles > 0 ? MessageBoxImage.Warning : MessageBoxImage.Error);
                                        }
                                        else
                                        {
                                            MessageBox.Show(
                                                message,
                                                "Success",
                                                MessageBoxButton.OK,
                                                MessageBoxImage.Information);
                                        }
                                    }
                                    catch (Exception uiEx)
                                    {
                                        // Fallback error handling
                                        System.Diagnostics.Debug.WriteLine($"[MainWindow] Error showing results: {uiEx.Message}");
                                        System.Diagnostics.Debug.WriteLine($"[MainWindow] StackTrace: {uiEx.StackTrace}");
                                        MessageBox.Show(
                                            $"Error displaying results: {uiEx.Message}",
                                            "Error",
                                            MessageBoxButton.OK,
                                            MessageBoxImage.Error);
                                    }
                                }
                            }), System.Windows.Threading.DispatcherPriority.Normal);
                            System.Diagnostics.Debug.WriteLine($"[MainWindow] Dispatcher.BeginInvoke for results called successfully");
                        }
                        else
                        {
                            // Fallback if Dispatcher is not available
                            System.Diagnostics.Debug.WriteLine("[MainWindow] Dispatcher not available for showing results");
                            System.Diagnostics.Debug.WriteLine($"[MainWindow] progressWindowRef is null: {progressWindowRef == null}");
                            System.Diagnostics.Debug.WriteLine($"[MainWindow] progressWindowRef.Dispatcher is null: {progressWindowRef?.Dispatcher == null}");
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"[MainWindow] Exception in Task.Run: {ex.Message}");
                        System.Diagnostics.Debug.WriteLine($"[MainWindow] Exception type: {ex.GetType().FullName}");
                        System.Diagnostics.Debug.WriteLine($"[MainWindow] StackTrace: {ex.StackTrace}");
                        
                        // Copy exception message to local variable before Dispatcher.BeginInvoke
                        var errorMessage = ex.Message ?? "Unknown error";
                        
                        System.Diagnostics.Debug.WriteLine($"[MainWindow] progressWindowRef is null: {progressWindowRef == null}");
                        System.Diagnostics.Debug.WriteLine($"[MainWindow] progressWindowRef.Dispatcher is null: {progressWindowRef?.Dispatcher == null}");
                        
                        // Use the window's Dispatcher for thread safety
                        if (progressWindowRef != null && progressWindowRef.Dispatcher != null)
                        {
                            System.Diagnostics.Debug.WriteLine($"[MainWindow] Calling Dispatcher.BeginInvoke for error on thread: {System.Threading.Thread.CurrentThread.ManagedThreadId}");
                            try
                            {
                                progressWindowRef.Dispatcher.BeginInvoke(new Action(() =>
                                {
                                    System.Diagnostics.Debug.WriteLine($"[MainWindow] Inside error Dispatcher callback on thread: {System.Threading.Thread.CurrentThread.ManagedThreadId}");
                                    try
                                    {
                                        if (progressWindowRef != null)
                                        {
                                            progressWindowRef.Close();
                                        }
                                        SubmitButton.IsEnabled = true;
                                        MessageBox.Show(
                                            $"Error during update: {errorMessage}",
                                            "Error",
                                            MessageBoxButton.OK,
                                            MessageBoxImage.Error);
                                    }
                                    catch (Exception innerEx)
                                    {
                                        System.Diagnostics.Debug.WriteLine($"[MainWindow] Error in error handler: {innerEx.Message}");
                                        System.Diagnostics.Debug.WriteLine($"[MainWindow] Inner exception StackTrace: {innerEx.StackTrace}");
                                        // Ignore if window is already closed
                                    }
                                }), System.Windows.Threading.DispatcherPriority.Normal);
                                System.Diagnostics.Debug.WriteLine($"[MainWindow] Dispatcher.BeginInvoke for error called successfully");
                            }
                            catch (Exception dispatcherEx)
                            {
                                System.Diagnostics.Debug.WriteLine($"[MainWindow] Error calling Dispatcher.BeginInvoke for error: {dispatcherEx.Message}");
                                System.Diagnostics.Debug.WriteLine($"[MainWindow] Dispatcher exception StackTrace: {dispatcherEx.StackTrace}");
                            }
                        }
                        else
                        {
                            System.Diagnostics.Debug.WriteLine("[MainWindow] Dispatcher not available for showing error");
                        }
                    }
                });
            }
        }
    }
}
