using System.Windows;

namespace CadFilesUpdater.Windows
{
    public partial class ProgressWindow : Window
    {
        public ProgressWindow()
        {
            InitializeComponent();
        }

        public void UpdateProgress(int processed, int total, string currentFile = null)
        {
            Dispatcher.Invoke(() =>
            {
                ProgressBar.Maximum = total;
                ProgressBar.Value = processed;
                
                if (!string.IsNullOrEmpty(currentFile))
                {
                    StatusText.Text = $"Processing: {System.IO.Path.GetFileName(currentFile)} ({processed}/{total})";
                }
                else
                {
                    StatusText.Text = $"Processed: {processed}/{total}";
                }
            });
        }
    }
}
