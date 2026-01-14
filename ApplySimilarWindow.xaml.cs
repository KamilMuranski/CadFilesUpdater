using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;

namespace CadFilesUpdater.Windows
{
    public enum ApplySimilarMode
    {
        SelectedFiles,
        AllFiles
    }

    public partial class ApplySimilarWindow : Window
    {
        public ApplySimilarMode Mode { get; private set; } = ApplySimilarMode.SelectedFiles;
        public string ValueText => ValueBox.Text;
        public string SelectedBlockName { get; private set; }
        public string SelectedAttributeTag { get; private set; }
        public List<string> SelectedFilePathsForApply { get; private set; } = new List<string>();

        private readonly List<string> _allFilePaths;
        private readonly Func<string, List<BlockAnalyzer.AttributeInstanceRow>> _instanceLoader;

        public ApplySimilarWindow(
            string initialValue,
            string initialBlockName,
            string initialAttributeTag,
            List<string> allFilePaths,
            List<string> defaultSelectedFilePaths,
            Func<string, List<BlockAnalyzer.AttributeInstanceRow>> instanceLoader)
        {
            InitializeComponent();
            ValueBox.Text = initialValue ?? "";
            ValueBox.SelectAll();
            ValueBox.Focus();

            SelectedBlockName = initialBlockName;
            SelectedAttributeTag = initialAttributeTag;

            _allFilePaths = allFilePaths ?? new List<string>();
            SelectedFilePathsForApply = (defaultSelectedFilePaths ?? new List<string>()).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
            _instanceLoader = instanceLoader;

            RefreshTargetLabels();
        }

        private void RefreshTargetLabels()
        {
            BlockNameText.Text = string.IsNullOrWhiteSpace(SelectedBlockName) ? "(not set)" : SelectedBlockName;
            AttributeTagText.Text = string.IsNullOrWhiteSpace(SelectedAttributeTag) ? "(not set)" : SelectedAttributeTag;
        }

        private void ChooseAttribute_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new ChooseAttributeWindow(
                _allFilePaths,
                SelectedFilePathsForApply,
                SelectedBlockName,
                SelectedAttributeTag,
                _instanceLoader)
            {
                Owner = this
            };

            if (dlg.ShowDialog() != true) return;

            SelectedFilePathsForApply = dlg.SelectedFilePaths.ToList();
            SelectedBlockName = dlg.SelectedBlockName;
            SelectedAttributeTag = dlg.SelectedAttributeTag;
            RefreshTargetLabels();
        }

        private void ApplySelected_Click(object sender, RoutedEventArgs e)
        {
            Mode = ApplySimilarMode.SelectedFiles;
            DialogResult = true;
            Close();
        }

        private void ApplyAll_Click(object sender, RoutedEventArgs e)
        {
            Mode = ApplySimilarMode.AllFiles;
            DialogResult = true;
            Close();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}

