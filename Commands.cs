using System;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Runtime;
using CadFilesUpdater.Windows;

namespace CadFilesUpdater
{
    public class Commands
    {
        [CommandMethod("CadFilesUpdater", "UpdateFiles", CommandFlags.Modal | CommandFlags.NoActionRecording)]
        public void UpdateFiles()
        {
            try
            {
                var mainWindow = new MainWindow();
                mainWindow.Show();
                mainWindow.Activate();
                mainWindow.Focus();
                mainWindow.Topmost = true;
                mainWindow.Topmost = false;
            }
            catch (System.Exception ex)
            {
                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage($"\nBłąd: {ex.Message}\n");
            }
        }
    }
}
