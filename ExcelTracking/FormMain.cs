#region using
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using ExcelTracking;
using Ookii.Dialogs.Wpf;
using SettingsManager;
#endregion

namespace ExcelTracking
{
    public partial class FormMain : Form
    {
        public FormMain()
        {
            InitializeComponent();
            LoadFormSettings();
        }
        
        //===================================================================================
        // *** MAIN TRACKING TAB
        private void button_Select_Master_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
                openFileDialog.Title = "Select Master Excel File";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = openFileDialog.FileName;
                    txtFilePath_MasterFile.Text = filePath;
                    TrackingInputData.txtFilePath_Master = filePath;
                    SaveSettings();
                }
            }
        }
        
        private void button_Select_InputData_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
                openFileDialog.Title = "Select Input Data Excel File";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = openFileDialog.FileName;
                    txtFilePath_InputDataFile.Text = filePath;
                    TrackingInputData.txtFilePath_InputData = filePath;
                    SaveSettings();
                }
            }
        }

        private void button_Select_InputRecordMaster_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
                openFileDialog.Title = "Select InputRecordMaster Excel File";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = openFileDialog.FileName;
                    txtFilePath_InputRecordMasterFile.Text = filePath;
                    TrackingInputData.txtFilePath_InputRecordMaster = filePath;
                    SaveSettings();
                }
            }
        }

        private async void button_Check_InputData_Drawing_Click(object sender, EventArgs e)
        {
            string buttonName = "Check InputData file - Drawing";

            DisableButton();
            labelStatus.Text = buttonName + " is running ...";

            await Task.Run(() => TrackingInputData.CheckInputDataFile_Drawing());

            EnableButton();
            labelStatus.Text = buttonName + " is done!";
        }

        private async void button_Check_InputData_Model_Click(object sender, EventArgs e)
        {
            string buttonName = "Check InputData file - Model";

            DisableButton();
            labelStatus.Text = buttonName + " is running ...";

            await Task.Run(() => TrackingInputData.CheckInputDataFile_Model());

            EnableButton();
            labelStatus.Text = buttonName + " is done!";
        }

        // Drawing - WORKING
        private async void button_Transfer_Working_Receive_Click(object sender, EventArgs e)
        {
            string buttonName = "Transfer WORKING_Receive to Master";

            DisableButton();

            await Task.Run(() => TrackingInputData.Transfer_WORKING_Receive_ToMaster());

            EnableButton();
            labelStatus.Text = buttonName + " is done!";
        }
        
        // Drawing - 1st
        private async void button_Transfer_RLMU_Receive_1st_Click(object sender, EventArgs e)
        {
            string buttonName = "Transfer RLMU_Receive_1st to Master";
            
            DisableButton();
            labelStatus.Text = buttonName + " is running ...";

            await Task.Run(() => TrackingInputData.Transfer_RLMU_Receive_1st_ToMaster());

            EnableButton();
            labelStatus.Text = buttonName + " is done!";
        }
        private async void button_Transfer_Drawing_Submit_1st_Click(object sender, EventArgs e)
        {
            string buttonName = "Transfer Drawing_Submit_1st to Master";

            DisableButton();
            labelStatus.Text = buttonName + " is running ...";

            await Task.Run(() => TrackingInputData.Transfer_Drawing_Submit_1st_ToMaster());

            EnableButton();
            labelStatus.Text = buttonName + " is done!";
        }
        private async void button_Transfer_Drawing_RFI_1st_Click(object sender, EventArgs e)
        {
            string buttonName = "Transfer Drawing_RFI_1st to Master";
            DisableButton();
            labelStatus.Text = buttonName + " is running ...";

            await Task.Run(() => TrackingInputData.Transfer_Drawing_RFI_1st_ToMaster());

            EnableButton();
            labelStatus.Text = buttonName + " is done!";
        }
        private async void button_Transfer_Drawing_Feedback_1st_Click(object sender, EventArgs e)
        {
            string buttonName = "Transfer Drawing_Feedback_1st to Master";

            DisableButton();
            labelStatus.Text = buttonName + " is running ...";

            await Task.Run(() => TrackingInputData.Transfer_Drawing_Feedback_1st_ToMaster());

            EnableButton();
            labelStatus.Text = buttonName + " is done!";
        }

        // Drawing - Final
        private async void button_Transfer_RLMU_Receive_Final_Click(object sender, EventArgs e)
        {
            string buttonName = "Transfer RLMU_Receive_Final to Master";

            DisableButton();
            labelStatus.Text = buttonName + " is running ...";

            await Task.Run(() => TrackingInputData.Transfer_RLMU_Receive_Final_ToMaster());

            EnableButton();
            labelStatus.Text = buttonName + " is done!";
        }
        private async void button_Transfer_Drawing_Submit_Final_Click(object sender, EventArgs e)
        {
            string buttonName = "Transfer Drawing_Submit_Final to Master";

            DisableButton();
            labelStatus.Text = buttonName + " is running ...";

            await Task.Run(() => TrackingInputData.Transfer_Drawing_Submit_Final_ToMaster());

            EnableButton();
            labelStatus.Text = buttonName + " is done!";
        }
        private async void button_Transfer_Drawing_RFI_Final_Click(object sender, EventArgs e)
        {
            string buttonName = "Transfer Drawing_RFI_Final to Master";
            DisableButton();
            labelStatus.Text = buttonName + " is running ...";

            await Task.Run(() => TrackingInputData.Transfer_Drawing_RFI_Final_ToMaster());

            EnableButton();
            labelStatus.Text = buttonName + " is done!";
        }
        private async void button_Transfer_Drawing_Feedback_Final_Click(object sender, EventArgs e)
        {
            string buttonName = "Transfer Drawing_Feedback_Final to Master";

            DisableButton();
            labelStatus.Text = buttonName + " is running ...";

            await Task.Run(() => TrackingInputData.Transfer_Drawing_Feedback_Final_ToMaster());

            EnableButton();
            labelStatus.Text = buttonName + " is done!";
        }

        // Model
        private async void button_Transfer_Model_Receive_Click(object sender, EventArgs e)
        {
            string buttonName = "Transfer Model_Receive to Master";

            DisableButton();
            labelStatus.Text = buttonName + " is running ...";

            await Task.Run(() => TrackingInputData.Transfer_Model_Receive_ToMaster());

            EnableButton();
            labelStatus.Text = buttonName + " is done!";
        }
        private async void button_Transfer_Model_Submit_Click(object sender, EventArgs e)
        {
            string buttonName = "Transfer Model_Submit to Master";

            DisableButton();
            labelStatus.Text = buttonName + " is running ...";

            await Task.Run(() => TrackingInputData.Transfer_Model_Submit_ToMaster());

            EnableButton();
            labelStatus.Text = buttonName + " is done!";
        }
        private async void button_Transfer_Model_Feedback_Click(object sender, EventArgs e)
        {
            string buttonName = "Transfer Model_Feedback to Master";

            DisableButton();
            labelStatus.Text = buttonName + " is running ...";

            await Task.Run(() => TrackingInputData.Transfer_Model_Feedback_ToMaster());

            EnableButton();
            labelStatus.Text = buttonName + " is done!";
        }

        //===================================================================================
        // *** TIMESHEET TAB
        private void button_ts_Select_Master_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
                openFileDialog.Title = "Select Master Excel File";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = openFileDialog.FileName;
                    txtFilePath_TS_MasterFile.Text = filePath;
                    TrackingInputData.txtFilePath_Master = filePath;
                }
            }
        }

        private void button_ts_Select_InputTS_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
                openFileDialog.Title = "Select Input Data Excel File";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = openFileDialog.FileName;
                    txtFilePath_TS_InputDataFile.Text = filePath;
                    TrackingInputData.txtFilePath_TS_InputData = filePath;
                }
            }
        }

        private void button_Select_TS_RecordMaster_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
                openFileDialog.Title = "Select InputRecordMaster Excel File";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = openFileDialog.FileName;
                    txtFilePath_TS_InputRecordMasterFile.Text = filePath;
                    TrackingInputData.txtFilePath_TS_InputRecordMaster = filePath;
                }
            }
        }

        private async void button_ts_Transfer_TS_Drawing_Click(object sender, EventArgs e)
        {
            string buttonName = "Transfer Drawing_TimeSheet to Master";

            DisableButton();
            labelStatus_TS.Text = buttonName + " is running ...";

            await Task.Run(() => TrackingInputData.Transfer_Drawing_TimeSheet());

            EnableButton();
            labelStatus_TS.Text = buttonName + " is done!";
        }

        private async void button_ts_Transfer_TS_Model_Click(object sender, EventArgs e)
        {
            string buttonName = "Transfer Model_TimeSheet to Master";

            DisableButton();
            labelStatus_TS.Text = buttonName + " is running ...";

            await Task.Run(() => TrackingInputData.Transfer_Model_TimeSheet());

            EnableButton();
            labelStatus_TS.Text = buttonName + " is done!";
        }

        //===================================================================================
        // *** OUTPUT TAB
        private void button_Output_Form_Click(object sender, EventArgs e)
        {
            TrackingInputData.ExportData_OutputForm();
        }

        //===================================================================================
        // *** Other Functions
        private void DisableButton()
        {
            // Main Tab
            button_Select_Master.Enabled = false;
            button_Select_InputData.Enabled = false;
            button_Select_InputRecordMaster.Enabled = false;

            button_Check_InputData_Drawing.Enabled = false;
            button_Check_InputData_Model.Enabled = false;

            button_Transfer_Working_Receive.Enabled = false;
            button_Transfer_RLMU_Receive_1st.Enabled = false;
            button_Transfer_Drawing_Submit_1st.Enabled = false;
            button_Transfer_Drawing_RFI_1st.Enabled = false;
            button_Transfer_Drawing_Feedback_1st.Enabled = false;

            button_Transfer_RLMU_Receive_Final.Enabled = false;
            button_Transfer_Drawing_Submit_Final.Enabled = false;
            button_Transfer_Drawing_RFI_Final.Enabled = false;
            button_Transfer_Drawing_Feedback_Final.Enabled = false;

            button_Transfer_Model_Receive.Enabled = false;
            button_Transfer_Model_Submit.Enabled = false;
            button_Transfer_Model_Feedback.Enabled = false;

            // TimeSheet Tab
            button_ts_Select_Master.Enabled = false;
            button_ts_Select_InputTS.Enabled = false;
            button_Select_TS_RecordMaster.Enabled = false;
            button_ts_Transfer_TS_Drawing.Enabled = false;
            button_ts_Transfer_TS_Model.Enabled = false;
        }

        private void EnableButton()
        {
            // Main Tab
            button_Select_Master.Enabled = true;
            button_Select_InputData.Enabled = true;
            button_Select_InputRecordMaster.Enabled = true;

            button_Check_InputData_Drawing.Enabled = true;
            button_Check_InputData_Model.Enabled = true;

            button_Transfer_Working_Receive.Enabled = true;
            button_Transfer_RLMU_Receive_1st.Enabled = true;
            button_Transfer_Drawing_Submit_1st.Enabled = true;
            button_Transfer_Drawing_RFI_1st.Enabled = true;
            button_Transfer_Drawing_Feedback_1st.Enabled = true;

            button_Transfer_RLMU_Receive_Final.Enabled = true;
            button_Transfer_Drawing_Submit_Final.Enabled = true;
            button_Transfer_Drawing_RFI_Final.Enabled = true;
            button_Transfer_Drawing_Feedback_Final.Enabled = true;

            button_Transfer_Model_Receive.Enabled = true;
            button_Transfer_Model_Submit.Enabled = true;
            button_Transfer_Model_Feedback.Enabled = true;

            // TimeSheet Tab
            button_ts_Select_Master.Enabled = true;
            button_ts_Select_InputTS.Enabled = true;
            button_Select_TS_RecordMaster.Enabled = true;
            button_ts_Transfer_TS_Drawing.Enabled = true;
            button_ts_Transfer_TS_Model.Enabled = true;

        }



        //==============================================================
        private FormSettings_MainTracking settings;
        // 📂 Load settings
        private void LoadFormSettings()
        {
            settings = SettingsManagerConfig.LoadSettings_MainTracking();

            // Apply vào textboxes (sẽ trống nếu lần đầu)
            txtFilePath_MasterFile.Text = settings.MasterFile;
            txtFilePath_InputDataFile.Text = settings.InputDataFile;
            txtFilePath_InputRecordMasterFile.Text = settings.RecordMasterFile;
            
            TrackingInputData.txtFilePath_Master = settings.MasterFile;
            TrackingInputData.txtFilePath_InputData = settings.InputDataFile;
            TrackingInputData.txtFilePath_InputRecordMaster = settings.RecordMasterFile;

        }
        // 💾 Save settings
        private void SaveSettings()
        {
            settings.MasterFile = txtFilePath_MasterFile.Text;
            settings.InputDataFile = txtFilePath_InputDataFile.Text;
            settings.RecordMasterFile = txtFilePath_InputRecordMasterFile.Text;

            SettingsManagerConfig.SaveSettings_MainTracking(settings);
        }

    }
}
