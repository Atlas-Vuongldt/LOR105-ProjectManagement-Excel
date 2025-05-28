#region references
using System;
using Ookii.Dialogs.Wpf;
#endregion

namespace LDTV
{
    public class PTAChecking
    {
        //===================================================================
        // Other function
        static string SelectFolder(string description)
        {
            #region
            VistaFolderBrowserDialog folderDialog = new VistaFolderBrowserDialog();
            folderDialog.ShowNewFolderButton = true;
            folderDialog.UseDescriptionForTitle = true;

            folderDialog.RootFolder = Environment.SpecialFolder.Desktop;
            folderDialog.Description = description;
            
            if (folderDialog.ShowDialog() == true)
            {
                return folderDialog.SelectedPath;
            }
            return null;
            #endregion
        }
    }
}