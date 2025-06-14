﻿#region
using Ookii.Dialogs.Wpf;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SettingsManager;
using System.IO;
#endregion

namespace ExcelTracking
{
    public partial class GetInputData : Form
    {
        public GetInputData()
        {
            InitializeComponent();
            LoadFormSettings();
        }

        //===================================================================================
        // *** GET FILES IN FOLDER TAB
        private void button_GetFile_Select_InputFolder_Click(object sender, EventArgs e)
        {
            VistaFolderBrowserDialog folderDialog = new VistaFolderBrowserDialog();
            folderDialog.ShowNewFolderButton = true;
            folderDialog.UseDescriptionForTitle = true;

            folderDialog.RootFolder = Environment.SpecialFolder.Desktop;
            folderDialog.Description = "Select Input Folder";

            if (folderDialog.ShowDialog() == true)
            {
                string folderPath = folderDialog.SelectedPath;
                txtFilePath_GetFile_InputDataFolder.Text = folderPath;
                TrackingInputData.txtFilePath_GetFile_InputDataFolder = folderPath;
            }
        }

        private void button_GetFile_Select_Master_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
                openFileDialog.Title = "Select Master Excel File";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = openFileDialog.FileName;
                    txtFilePath_GetFile_MasterFile.Text = filePath;
                    TrackingInputData.txtFilePath_Master = filePath;
                }
            }
        }

        private void button_GetFile_Select_OutputFolder_Click(object sender, EventArgs e)
        {
            VistaFolderBrowserDialog folderDialog = new VistaFolderBrowserDialog();
            folderDialog.ShowNewFolderButton = true;
            folderDialog.UseDescriptionForTitle = true;

            folderDialog.RootFolder = Environment.SpecialFolder.Desktop;
            folderDialog.Description = "Select Output Folder";

            if (folderDialog.ShowDialog() == true)
            {
                string folderPath = folderDialog.SelectedPath;
                txtFilePath_GetFile_OutputDataFolder.Text = folderPath;
                TrackingInputData.txtFilePath_GetFile_OutputDataFolder = folderPath;
            }
        }

        private async void button_GetFileNameFromInputFolder_Click(object sender, EventArgs e)
        {
            SaveSettings();

            DisableButton();

            TrackingInputData.isInputAsPTANo = IsInputAsPTANo.Checked;
            TrackingInputData.isInputAsModel = IsInputAsModel.Checked;
            TrackingInputData.isGetCADInfo = false;
            await Task.Run(() => TrackingInputData.GetFileName_FromInputFolder());

            EnableButton();
        }

        private void button_Output_Form_Click(object sender, EventArgs e)
        {
            TrackingInputData.ExportData_OutputForm();
        }
        private void DisableButton()
        {
            // GetFilesInFolder Tab
            button_GetFile_Select_InputFolder.Enabled = false;
            button_GetFile_Select_OutputFolder.Enabled = false;
            button_GetFile_Select_Master.Enabled = false;
            button_GetFileNameFromInputFolder.Enabled = false;
            IsInputAsPTANo.Enabled = false;
            IsInputAsModel.Enabled = false;
        }

        private void EnableButton()
        {
            // GetFilesInFolder Tab
            button_GetFile_Select_InputFolder.Enabled = true;
            button_GetFile_Select_OutputFolder.Enabled = true;
            button_GetFile_Select_Master.Enabled = true;
            button_GetFileNameFromInputFolder.Enabled = true;
            IsInputAsPTANo.Enabled = true;
            IsInputAsModel.Enabled = true;
        }


        //==============================================================
        private FormSettings_GetInputData settings;
        // 📂 Load settings
        private void LoadFormSettings()
        {
            settings = SettingsManagerConfig.LoadSettings_GetInputData();

            // Apply vào textboxes (sẽ trống nếu lần đầu)
            txtFilePath_GetFile_InputDataFolder.Text = settings.InputFolder;
            txtFilePath_GetFile_OutputDataFolder.Text = settings.OutputFolder;
            txtFilePath_GetFile_MasterFile.Text = settings.MasterFile;
        }
        // 💾 Save settings
        private void SaveSettings()
        {
            settings.InputFolder = txtFilePath_GetFile_InputDataFolder.Text;
            settings.OutputFolder = txtFilePath_GetFile_OutputDataFolder.Text;
            settings.MasterFile = txtFilePath_GetFile_MasterFile.Text;

            SettingsManagerConfig.SaveSettings_GetInputData(settings);
        }

        private void Open_CadConfigFolder_Click(object sender, EventArgs e)
        {
            string folderPath = CadInfoExtractor.settingFolder;
            try
            {
                if (string.IsNullOrEmpty(folderPath))
                {
                    MessageBox.Show("❌ Folder path is empty!", "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (!Directory.Exists(folderPath))
                {
                    MessageBox.Show($"❌ Folder not found:\n{folderPath}", "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                System.Diagnostics.Process.Start("explorer.exe", folderPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ Error opening folder:\n{ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }






    }
}
