using System;
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
            // This method is now called from UI thread (via MainWindow's Dispatcher)
            // So we can directly update UI elements
            try
            {
                if (ProgressBar != null)
                {
                    ProgressBar.Maximum = total;
                    ProgressBar.Value = processed;
                }
                
                if (StatusText != null)
                {
                    if (!string.IsNullOrEmpty(currentFile))
                    {
                        StatusText.Text = $"Processing: {System.IO.Path.GetFileName(currentFile)} ({processed}/{total})";
                    }
                    else
                    {
                        StatusText.Text = $"Processed: {processed}/{total}";
                    }
                }
            }
            catch
            {
                // Ignore errors if window is being closed
            }
        }
    }
}
