using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Runtime;
using CadFilesUpdater.Windows;

namespace CadFilesUpdater
{
    public class Commands
    {
        [CommandMethod("CadFilesUpdater", "UpdateAttributesInFiles", CommandFlags.Modal | CommandFlags.NoActionRecording)]
        public void UpdateAttributesInFiles()
        {
            try
            {
                var mainWindow = new MainWindow();
                
                // Show window first
                mainWindow.Show();
                mainWindow.WindowState = WindowState.Normal;
                
                // Force window to front using Windows API
                var helper = new WindowInteropHelper(mainWindow);
                helper.EnsureHandle();
                // Use HWND_TOP to bring to front (not permanently forcing, but ensures visibility)
                SetWindowPos(helper.Handle, new IntPtr(-1), 0, 0, 0, 0, 0x0001 | 0x0002); // SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW
                
                // Activate and focus
                mainWindow.Activate();
                mainWindow.Focus();
                mainWindow.BringIntoView();
            }
            catch (System.Exception ex)
            {
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage($"\nError: {ex.Message}\n");
            }
        }
        
        [DllImport("user32.dll")]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
    }
}
