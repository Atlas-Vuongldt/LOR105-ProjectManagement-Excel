
namespace ExcelTracking
{
    partial class FormMain
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormMain));
            this.button_Select_Master = new System.Windows.Forms.Button();
            this.button_Check_InputData_Drawing = new System.Windows.Forms.Button();
            this.txtFilePath_MasterFile = new System.Windows.Forms.TextBox();
            this.button_Select_InputData = new System.Windows.Forms.Button();
            this.txtFilePath_InputDataFile = new System.Windows.Forms.TextBox();
            this.button_Select_InputRecordMaster = new System.Windows.Forms.Button();
            this.txtFilePath_InputRecordMasterFile = new System.Windows.Forms.TextBox();
            this.button_Transfer_RLMU_Receive_1st = new System.Windows.Forms.Button();
            this.button_Transfer_RLMU_Receive_Final = new System.Windows.Forms.Button();
            this.button_Transfer_Drawing_Submit_1st = new System.Windows.Forms.Button();
            this.button_Transfer_Drawing_Submit_Final = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.button_Transfer_Working_Receive = new System.Windows.Forms.Button();
            this.button_Transfer_Drawing_Feedback_Final = new System.Windows.Forms.Button();
            this.button_Transfer_Drawing_Feedback_1st = new System.Windows.Forms.Button();
            this.button_Transfer_Drawing_RFI_Final = new System.Windows.Forms.Button();
            this.button_Transfer_Drawing_RFI_1st = new System.Windows.Forms.Button();
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.button_Transfer_Model_Feedback = new System.Windows.Forms.Button();
            this.button_Transfer_Model_Receive = new System.Windows.Forms.Button();
            this.button_Transfer_Model_Submit = new System.Windows.Forms.Button();
            this.labelStatus = new System.Windows.Forms.Label();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.button_Check_InputData_Model = new System.Windows.Forms.Button();
            this.tableLayoutPanel2 = new System.Windows.Forms.TableLayoutPanel();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.labelStatus_TS = new System.Windows.Forms.Label();
            this.txtFilePath_TS_InputRecordMasterFile = new System.Windows.Forms.TextBox();
            this.txtFilePath_TS_InputDataFile = new System.Windows.Forms.TextBox();
            this.txtFilePath_TS_MasterFile = new System.Windows.Forms.TextBox();
            this.button_Select_TS_RecordMaster = new System.Windows.Forms.Button();
            this.button_ts_Transfer_TS_Model = new System.Windows.Forms.Button();
            this.button_ts_Transfer_TS_Drawing = new System.Windows.Forms.Button();
            this.button_ts_Select_InputTS = new System.Windows.Forms.Button();
            this.button_ts_Select_Master = new System.Windows.Forms.Button();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.button_Output_Form = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tableLayoutPanel2.SuspendLayout();
            this.tableLayoutPanel1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.tabPage3.SuspendLayout();
            this.SuspendLayout();
            // 
            // button_Select_Master
            // 
            this.button_Select_Master.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.button_Select_Master.Dock = System.Windows.Forms.DockStyle.Top;
            this.button_Select_Master.Location = new System.Drawing.Point(3, 23);
            this.button_Select_Master.Name = "button_Select_Master";
            this.button_Select_Master.Size = new System.Drawing.Size(144, 44);
            this.button_Select_Master.TabIndex = 0;
            this.button_Select_Master.Text = "Select Master file";
            this.button_Select_Master.UseVisualStyleBackColor = true;
            this.button_Select_Master.Click += new System.EventHandler(this.button_Select_Master_Click);
            // 
            // button_Check_InputData_Drawing
            // 
            this.button_Check_InputData_Drawing.BackColor = System.Drawing.SystemColors.InactiveCaption;
            this.button_Check_InputData_Drawing.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.button_Check_InputData_Drawing.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.button_Check_InputData_Drawing.Location = new System.Drawing.Point(6, 250);
            this.button_Check_InputData_Drawing.Name = "button_Check_InputData_Drawing";
            this.button_Check_InputData_Drawing.Size = new System.Drawing.Size(230, 38);
            this.button_Check_InputData_Drawing.TabIndex = 1;
            this.button_Check_InputData_Drawing.Text = "Check InputData File - Drawing";
            this.button_Check_InputData_Drawing.UseVisualStyleBackColor = false;
            this.button_Check_InputData_Drawing.Click += new System.EventHandler(this.button_Check_InputData_Drawing_Click);
            // 
            // txtFilePath_MasterFile
            // 
            this.txtFilePath_MasterFile.BackColor = System.Drawing.Color.White;
            this.txtFilePath_MasterFile.Dock = System.Windows.Forms.DockStyle.Top;
            this.txtFilePath_MasterFile.ForeColor = System.Drawing.Color.Red;
            this.txtFilePath_MasterFile.Location = new System.Drawing.Point(153, 23);
            this.txtFilePath_MasterFile.Multiline = true;
            this.txtFilePath_MasterFile.Name = "txtFilePath_MasterFile";
            this.txtFilePath_MasterFile.ReadOnly = true;
            this.txtFilePath_MasterFile.Size = new System.Drawing.Size(725, 44);
            this.txtFilePath_MasterFile.TabIndex = 4;
            // 
            // button_Select_InputData
            // 
            this.button_Select_InputData.Dock = System.Windows.Forms.DockStyle.Top;
            this.button_Select_InputData.Location = new System.Drawing.Point(3, 73);
            this.button_Select_InputData.Name = "button_Select_InputData";
            this.button_Select_InputData.Size = new System.Drawing.Size(144, 44);
            this.button_Select_InputData.TabIndex = 5;
            this.button_Select_InputData.Text = "Select InputData file";
            this.button_Select_InputData.UseVisualStyleBackColor = true;
            this.button_Select_InputData.Click += new System.EventHandler(this.button_Select_InputData_Click);
            // 
            // txtFilePath_InputDataFile
            // 
            this.txtFilePath_InputDataFile.BackColor = System.Drawing.Color.White;
            this.txtFilePath_InputDataFile.Dock = System.Windows.Forms.DockStyle.Top;
            this.txtFilePath_InputDataFile.ForeColor = System.Drawing.Color.Red;
            this.txtFilePath_InputDataFile.Location = new System.Drawing.Point(153, 73);
            this.txtFilePath_InputDataFile.Multiline = true;
            this.txtFilePath_InputDataFile.Name = "txtFilePath_InputDataFile";
            this.txtFilePath_InputDataFile.ReadOnly = true;
            this.txtFilePath_InputDataFile.Size = new System.Drawing.Size(725, 44);
            this.txtFilePath_InputDataFile.TabIndex = 6;
            // 
            // button_Select_InputRecordMaster
            // 
            this.button_Select_InputRecordMaster.Dock = System.Windows.Forms.DockStyle.Top;
            this.button_Select_InputRecordMaster.Location = new System.Drawing.Point(3, 123);
            this.button_Select_InputRecordMaster.Name = "button_Select_InputRecordMaster";
            this.button_Select_InputRecordMaster.Size = new System.Drawing.Size(144, 45);
            this.button_Select_InputRecordMaster.TabIndex = 8;
            this.button_Select_InputRecordMaster.Text = "Select InputRecordMaster file";
            this.button_Select_InputRecordMaster.UseVisualStyleBackColor = true;
            this.button_Select_InputRecordMaster.Click += new System.EventHandler(this.button_Select_InputRecordMaster_Click);
            // 
            // txtFilePath_InputRecordMasterFile
            // 
            this.txtFilePath_InputRecordMasterFile.BackColor = System.Drawing.Color.White;
            this.txtFilePath_InputRecordMasterFile.Dock = System.Windows.Forms.DockStyle.Top;
            this.txtFilePath_InputRecordMasterFile.ForeColor = System.Drawing.Color.Red;
            this.txtFilePath_InputRecordMasterFile.Location = new System.Drawing.Point(153, 123);
            this.txtFilePath_InputRecordMasterFile.Multiline = true;
            this.txtFilePath_InputRecordMasterFile.Name = "txtFilePath_InputRecordMasterFile";
            this.txtFilePath_InputRecordMasterFile.ReadOnly = true;
            this.txtFilePath_InputRecordMasterFile.Size = new System.Drawing.Size(725, 45);
            this.txtFilePath_InputRecordMasterFile.TabIndex = 9;
            // 
            // button_Transfer_RLMU_Receive_1st
            // 
            this.button_Transfer_RLMU_Receive_1st.BackColor = System.Drawing.SystemColors.InactiveCaption;
            this.button_Transfer_RLMU_Receive_1st.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.button_Transfer_RLMU_Receive_1st.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.button_Transfer_RLMU_Receive_1st.Location = new System.Drawing.Point(5, 65);
            this.button_Transfer_RLMU_Receive_1st.Name = "button_Transfer_RLMU_Receive_1st";
            this.button_Transfer_RLMU_Receive_1st.Size = new System.Drawing.Size(210, 40);
            this.button_Transfer_RLMU_Receive_1st.TabIndex = 13;
            this.button_Transfer_RLMU_Receive_1st.Text = "Transfer RLMU_Receive_1st to Master";
            this.button_Transfer_RLMU_Receive_1st.UseVisualStyleBackColor = false;
            this.button_Transfer_RLMU_Receive_1st.Click += new System.EventHandler(this.button_Transfer_RLMU_Receive_1st_Click);
            // 
            // button_Transfer_RLMU_Receive_Final
            // 
            this.button_Transfer_RLMU_Receive_Final.BackColor = System.Drawing.SystemColors.InactiveCaption;
            this.button_Transfer_RLMU_Receive_Final.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.button_Transfer_RLMU_Receive_Final.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.button_Transfer_RLMU_Receive_Final.Location = new System.Drawing.Point(5, 113);
            this.button_Transfer_RLMU_Receive_Final.Name = "button_Transfer_RLMU_Receive_Final";
            this.button_Transfer_RLMU_Receive_Final.Size = new System.Drawing.Size(210, 40);
            this.button_Transfer_RLMU_Receive_Final.TabIndex = 14;
            this.button_Transfer_RLMU_Receive_Final.Text = "Transfer RLMU_Receive_Final to Master";
            this.button_Transfer_RLMU_Receive_Final.UseVisualStyleBackColor = false;
            this.button_Transfer_RLMU_Receive_Final.Click += new System.EventHandler(this.button_Transfer_RLMU_Receive_Final_Click);
            // 
            // button_Transfer_Drawing_Submit_1st
            // 
            this.button_Transfer_Drawing_Submit_1st.BackColor = System.Drawing.SystemColors.InactiveCaption;
            this.button_Transfer_Drawing_Submit_1st.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.button_Transfer_Drawing_Submit_1st.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.button_Transfer_Drawing_Submit_1st.Location = new System.Drawing.Point(223, 65);
            this.button_Transfer_Drawing_Submit_1st.Name = "button_Transfer_Drawing_Submit_1st";
            this.button_Transfer_Drawing_Submit_1st.Size = new System.Drawing.Size(210, 40);
            this.button_Transfer_Drawing_Submit_1st.TabIndex = 15;
            this.button_Transfer_Drawing_Submit_1st.Text = "Transfer Drawing_Submit_1st to Master";
            this.button_Transfer_Drawing_Submit_1st.UseVisualStyleBackColor = false;
            this.button_Transfer_Drawing_Submit_1st.Click += new System.EventHandler(this.button_Transfer_Drawing_Submit_1st_Click);
            // 
            // button_Transfer_Drawing_Submit_Final
            // 
            this.button_Transfer_Drawing_Submit_Final.BackColor = System.Drawing.SystemColors.InactiveCaption;
            this.button_Transfer_Drawing_Submit_Final.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.button_Transfer_Drawing_Submit_Final.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.button_Transfer_Drawing_Submit_Final.Location = new System.Drawing.Point(223, 113);
            this.button_Transfer_Drawing_Submit_Final.Name = "button_Transfer_Drawing_Submit_Final";
            this.button_Transfer_Drawing_Submit_Final.Size = new System.Drawing.Size(210, 40);
            this.button_Transfer_Drawing_Submit_Final.TabIndex = 16;
            this.button_Transfer_Drawing_Submit_Final.Text = "Transfer Drawing_Submit_Final to Master";
            this.button_Transfer_Drawing_Submit_Final.UseVisualStyleBackColor = false;
            this.button_Transfer_Drawing_Submit_Final.Click += new System.EventHandler(this.button_Transfer_Drawing_Submit_Final_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.BackColor = System.Drawing.SystemColors.GradientInactiveCaption;
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.button_Transfer_Working_Receive);
            this.groupBox1.Controls.Add(this.button_Transfer_Drawing_Feedback_Final);
            this.groupBox1.Controls.Add(this.button_Transfer_Drawing_Feedback_1st);
            this.groupBox1.Controls.Add(this.button_Transfer_RLMU_Receive_1st);
            this.groupBox1.Controls.Add(this.button_Transfer_Drawing_Submit_Final);
            this.groupBox1.Controls.Add(this.button_Transfer_RLMU_Receive_Final);
            this.groupBox1.Controls.Add(this.button_Transfer_Drawing_Submit_1st);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Top;
            this.groupBox1.Location = new System.Drawing.Point(3, 3);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(875, 159);
            this.groupBox1.TabIndex = 17;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "DRAWING";
            // 
            // button_Transfer_Working_Receive
            // 
            this.button_Transfer_Working_Receive.BackColor = System.Drawing.SystemColors.InactiveCaption;
            this.button_Transfer_Working_Receive.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.button_Transfer_Working_Receive.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.button_Transfer_Working_Receive.Location = new System.Drawing.Point(6, 19);
            this.button_Transfer_Working_Receive.Name = "button_Transfer_Working_Receive";
            this.button_Transfer_Working_Receive.Size = new System.Drawing.Size(210, 40);
            this.button_Transfer_Working_Receive.TabIndex = 19;
            this.button_Transfer_Working_Receive.Text = "Transfer Working_Receive to Master";
            this.button_Transfer_Working_Receive.UseVisualStyleBackColor = false;
            this.button_Transfer_Working_Receive.Click += new System.EventHandler(this.button_Transfer_Working_Receive_Click);
            // 
            // button_Transfer_Drawing_Feedback_Final
            // 
            this.button_Transfer_Drawing_Feedback_Final.BackColor = System.Drawing.SystemColors.InactiveCaption;
            this.button_Transfer_Drawing_Feedback_Final.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.button_Transfer_Drawing_Feedback_Final.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.button_Transfer_Drawing_Feedback_Final.Location = new System.Drawing.Point(442, 113);
            this.button_Transfer_Drawing_Feedback_Final.Name = "button_Transfer_Drawing_Feedback_Final";
            this.button_Transfer_Drawing_Feedback_Final.Size = new System.Drawing.Size(210, 40);
            this.button_Transfer_Drawing_Feedback_Final.TabIndex = 18;
            this.button_Transfer_Drawing_Feedback_Final.Text = "Transfer Drawing_Feedback_Final to Master";
            this.button_Transfer_Drawing_Feedback_Final.UseVisualStyleBackColor = false;
            this.button_Transfer_Drawing_Feedback_Final.Click += new System.EventHandler(this.button_Transfer_Drawing_Feedback_Final_Click);
            // 
            // button_Transfer_Drawing_Feedback_1st
            // 
            this.button_Transfer_Drawing_Feedback_1st.BackColor = System.Drawing.SystemColors.InactiveCaption;
            this.button_Transfer_Drawing_Feedback_1st.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.button_Transfer_Drawing_Feedback_1st.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.button_Transfer_Drawing_Feedback_1st.Location = new System.Drawing.Point(442, 65);
            this.button_Transfer_Drawing_Feedback_1st.Name = "button_Transfer_Drawing_Feedback_1st";
            this.button_Transfer_Drawing_Feedback_1st.Size = new System.Drawing.Size(210, 40);
            this.button_Transfer_Drawing_Feedback_1st.TabIndex = 17;
            this.button_Transfer_Drawing_Feedback_1st.Text = "Transfer Drawing_Feedback_1st to Master";
            this.button_Transfer_Drawing_Feedback_1st.UseVisualStyleBackColor = false;
            this.button_Transfer_Drawing_Feedback_1st.Click += new System.EventHandler(this.button_Transfer_Drawing_Feedback_1st_Click);
            // 
            // button_Transfer_Drawing_RFI_Final
            // 
            this.button_Transfer_Drawing_RFI_Final.BackColor = System.Drawing.SystemColors.InactiveCaption;
            this.button_Transfer_Drawing_RFI_Final.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.button_Transfer_Drawing_RFI_Final.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.button_Transfer_Drawing_RFI_Final.Location = new System.Drawing.Point(671, 256);
            this.button_Transfer_Drawing_RFI_Final.Name = "button_Transfer_Drawing_RFI_Final";
            this.button_Transfer_Drawing_RFI_Final.Size = new System.Drawing.Size(210, 40);
            this.button_Transfer_Drawing_RFI_Final.TabIndex = 20;
            this.button_Transfer_Drawing_RFI_Final.Text = "Transfer Drawing_RFI_Final to Master";
            this.button_Transfer_Drawing_RFI_Final.UseVisualStyleBackColor = false;
            this.button_Transfer_Drawing_RFI_Final.Visible = false;
            this.button_Transfer_Drawing_RFI_Final.Click += new System.EventHandler(this.button_Transfer_Drawing_RFI_Final_Click);
            // 
            // button_Transfer_Drawing_RFI_1st
            // 
            this.button_Transfer_Drawing_RFI_1st.BackColor = System.Drawing.SystemColors.InactiveCaption;
            this.button_Transfer_Drawing_RFI_1st.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.button_Transfer_Drawing_RFI_1st.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.button_Transfer_Drawing_RFI_1st.Location = new System.Drawing.Point(671, 239);
            this.button_Transfer_Drawing_RFI_1st.Name = "button_Transfer_Drawing_RFI_1st";
            this.button_Transfer_Drawing_RFI_1st.Size = new System.Drawing.Size(210, 40);
            this.button_Transfer_Drawing_RFI_1st.TabIndex = 19;
            this.button_Transfer_Drawing_RFI_1st.Text = "Transfer Drawing_RFI_1st to Master";
            this.button_Transfer_Drawing_RFI_1st.UseVisualStyleBackColor = false;
            this.button_Transfer_Drawing_RFI_1st.Visible = false;
            this.button_Transfer_Drawing_RFI_1st.Click += new System.EventHandler(this.button_Transfer_Drawing_RFI_1st_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.BackColor = System.Drawing.SystemColors.GradientInactiveCaption;
            this.groupBox2.Controls.Add(this.button_Transfer_Model_Feedback);
            this.groupBox2.Controls.Add(this.button_Transfer_Model_Receive);
            this.groupBox2.Controls.Add(this.button_Transfer_Model_Submit);
            this.groupBox2.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.groupBox2.Location = new System.Drawing.Point(3, 196);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(875, 81);
            this.groupBox2.TabIndex = 18;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "MODEL";
            // 
            // button_Transfer_Model_Feedback
            // 
            this.button_Transfer_Model_Feedback.BackColor = System.Drawing.SystemColors.InactiveCaption;
            this.button_Transfer_Model_Feedback.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.button_Transfer_Model_Feedback.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.button_Transfer_Model_Feedback.Location = new System.Drawing.Point(442, 30);
            this.button_Transfer_Model_Feedback.Name = "button_Transfer_Model_Feedback";
            this.button_Transfer_Model_Feedback.Size = new System.Drawing.Size(210, 40);
            this.button_Transfer_Model_Feedback.TabIndex = 16;
            this.button_Transfer_Model_Feedback.Text = "Transfer Model_Feedback to Master";
            this.button_Transfer_Model_Feedback.UseVisualStyleBackColor = false;
            this.button_Transfer_Model_Feedback.Click += new System.EventHandler(this.button_Transfer_Model_Feedback_Click);
            // 
            // button_Transfer_Model_Receive
            // 
            this.button_Transfer_Model_Receive.BackColor = System.Drawing.SystemColors.InactiveCaption;
            this.button_Transfer_Model_Receive.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.button_Transfer_Model_Receive.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.button_Transfer_Model_Receive.Location = new System.Drawing.Point(5, 30);
            this.button_Transfer_Model_Receive.Name = "button_Transfer_Model_Receive";
            this.button_Transfer_Model_Receive.Size = new System.Drawing.Size(210, 40);
            this.button_Transfer_Model_Receive.TabIndex = 13;
            this.button_Transfer_Model_Receive.Text = "Transfer Model_Receive to Master";
            this.button_Transfer_Model_Receive.UseVisualStyleBackColor = false;
            this.button_Transfer_Model_Receive.Click += new System.EventHandler(this.button_Transfer_Model_Receive_Click);
            // 
            // button_Transfer_Model_Submit
            // 
            this.button_Transfer_Model_Submit.BackColor = System.Drawing.SystemColors.InactiveCaption;
            this.button_Transfer_Model_Submit.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.button_Transfer_Model_Submit.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.button_Transfer_Model_Submit.Location = new System.Drawing.Point(223, 30);
            this.button_Transfer_Model_Submit.Name = "button_Transfer_Model_Submit";
            this.button_Transfer_Model_Submit.Size = new System.Drawing.Size(210, 40);
            this.button_Transfer_Model_Submit.TabIndex = 15;
            this.button_Transfer_Model_Submit.Text = "Transfer Model_Submit to Master";
            this.button_Transfer_Model_Submit.UseVisualStyleBackColor = false;
            this.button_Transfer_Model_Submit.Click += new System.EventHandler(this.button_Transfer_Model_Submit_Click);
            // 
            // labelStatus
            // 
            this.labelStatus.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelStatus.ForeColor = System.Drawing.Color.Red;
            this.labelStatus.Location = new System.Drawing.Point(3, 183);
            this.labelStatus.Name = "labelStatus";
            this.labelStatus.Size = new System.Drawing.Size(881, 35);
            this.labelStatus.TabIndex = 19;
            this.labelStatus.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Controls.Add(this.tabPage3);
            this.tabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl1.ItemSize = new System.Drawing.Size(80, 25);
            this.tabControl1.Location = new System.Drawing.Point(0, 0);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(895, 621);
            this.tabControl1.TabIndex = 20;
            // 
            // tabPage1
            // 
            this.tabPage1.BackColor = System.Drawing.Color.Bisque;
            this.tabPage1.Controls.Add(this.button_Transfer_Drawing_RFI_Final);
            this.tabPage1.Controls.Add(this.button_Check_InputData_Model);
            this.tabPage1.Controls.Add(this.button_Transfer_Drawing_RFI_1st);
            this.tabPage1.Controls.Add(this.tableLayoutPanel2);
            this.tabPage1.Controls.Add(this.labelStatus);
            this.tabPage1.Controls.Add(this.tableLayoutPanel1);
            this.tabPage1.Controls.Add(this.button_Check_InputData_Drawing);
            this.tabPage1.Location = new System.Drawing.Point(4, 29);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(887, 588);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Main Tracking";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // button_Check_InputData_Model
            // 
            this.button_Check_InputData_Model.BackColor = System.Drawing.SystemColors.InactiveCaption;
            this.button_Check_InputData_Model.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.button_Check_InputData_Model.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.button_Check_InputData_Model.Location = new System.Drawing.Point(251, 250);
            this.button_Check_InputData_Model.Name = "button_Check_InputData_Model";
            this.button_Check_InputData_Model.Size = new System.Drawing.Size(230, 38);
            this.button_Check_InputData_Model.TabIndex = 21;
            this.button_Check_InputData_Model.Text = "Check InputData File - Model";
            this.button_Check_InputData_Model.UseVisualStyleBackColor = false;
            this.button_Check_InputData_Model.Click += new System.EventHandler(this.button_Check_InputData_Model_Click);
            // 
            // tableLayoutPanel2
            // 
            this.tableLayoutPanel2.ColumnCount = 1;
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel2.Controls.Add(this.groupBox1, 0, 0);
            this.tableLayoutPanel2.Controls.Add(this.groupBox2, 0, 1);
            this.tableLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.tableLayoutPanel2.Location = new System.Drawing.Point(3, 305);
            this.tableLayoutPanel2.Name = "tableLayoutPanel2";
            this.tableLayoutPanel2.RowCount = 2;
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 170F));
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 80F));
            this.tableLayoutPanel2.Size = new System.Drawing.Size(881, 280);
            this.tableLayoutPanel2.TabIndex = 20;
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 2;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 150F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.Controls.Add(this.button_Select_InputRecordMaster, 0, 3);
            this.tableLayoutPanel1.Controls.Add(this.txtFilePath_InputRecordMasterFile, 1, 3);
            this.tableLayoutPanel1.Controls.Add(this.button_Select_InputData, 0, 2);
            this.tableLayoutPanel1.Controls.Add(this.txtFilePath_InputDataFile, 1, 2);
            this.tableLayoutPanel1.Controls.Add(this.txtFilePath_MasterFile, 1, 1);
            this.tableLayoutPanel1.Controls.Add(this.button_Select_Master, 0, 1);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(3, 3);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 4;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 50F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 50F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 50F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(881, 177);
            this.tableLayoutPanel1.TabIndex = 19;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.labelStatus_TS);
            this.tabPage2.Controls.Add(this.txtFilePath_TS_InputRecordMasterFile);
            this.tabPage2.Controls.Add(this.txtFilePath_TS_InputDataFile);
            this.tabPage2.Controls.Add(this.txtFilePath_TS_MasterFile);
            this.tabPage2.Controls.Add(this.button_Select_TS_RecordMaster);
            this.tabPage2.Controls.Add(this.button_ts_Transfer_TS_Model);
            this.tabPage2.Controls.Add(this.button_ts_Transfer_TS_Drawing);
            this.tabPage2.Controls.Add(this.button_ts_Select_InputTS);
            this.tabPage2.Controls.Add(this.button_ts_Select_Master);
            this.tabPage2.Location = new System.Drawing.Point(4, 29);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(887, 588);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "TimeSheet";
            // 
            // labelStatus_TS
            // 
            this.labelStatus_TS.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelStatus_TS.ForeColor = System.Drawing.Color.Red;
            this.labelStatus_TS.Location = new System.Drawing.Point(6, 180);
            this.labelStatus_TS.Name = "labelStatus_TS";
            this.labelStatus_TS.Size = new System.Drawing.Size(800, 35);
            this.labelStatus_TS.TabIndex = 20;
            this.labelStatus_TS.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtFilePath_TS_InputRecordMasterFile
            // 
            this.txtFilePath_TS_InputRecordMasterFile.BackColor = System.Drawing.Color.White;
            this.txtFilePath_TS_InputRecordMasterFile.ForeColor = System.Drawing.Color.Red;
            this.txtFilePath_TS_InputRecordMasterFile.Location = new System.Drawing.Point(170, 121);
            this.txtFilePath_TS_InputRecordMasterFile.Multiline = true;
            this.txtFilePath_TS_InputRecordMasterFile.Name = "txtFilePath_TS_InputRecordMasterFile";
            this.txtFilePath_TS_InputRecordMasterFile.ReadOnly = true;
            this.txtFilePath_TS_InputRecordMasterFile.Size = new System.Drawing.Size(717, 44);
            this.txtFilePath_TS_InputRecordMasterFile.TabIndex = 8;
            // 
            // txtFilePath_TS_InputDataFile
            // 
            this.txtFilePath_TS_InputDataFile.BackColor = System.Drawing.Color.White;
            this.txtFilePath_TS_InputDataFile.ForeColor = System.Drawing.Color.Red;
            this.txtFilePath_TS_InputDataFile.Location = new System.Drawing.Point(170, 71);
            this.txtFilePath_TS_InputDataFile.Multiline = true;
            this.txtFilePath_TS_InputDataFile.Name = "txtFilePath_TS_InputDataFile";
            this.txtFilePath_TS_InputDataFile.ReadOnly = true;
            this.txtFilePath_TS_InputDataFile.Size = new System.Drawing.Size(717, 44);
            this.txtFilePath_TS_InputDataFile.TabIndex = 7;
            // 
            // txtFilePath_TS_MasterFile
            // 
            this.txtFilePath_TS_MasterFile.BackColor = System.Drawing.Color.White;
            this.txtFilePath_TS_MasterFile.ForeColor = System.Drawing.Color.Red;
            this.txtFilePath_TS_MasterFile.Location = new System.Drawing.Point(170, 21);
            this.txtFilePath_TS_MasterFile.Multiline = true;
            this.txtFilePath_TS_MasterFile.Name = "txtFilePath_TS_MasterFile";
            this.txtFilePath_TS_MasterFile.ReadOnly = true;
            this.txtFilePath_TS_MasterFile.Size = new System.Drawing.Size(717, 44);
            this.txtFilePath_TS_MasterFile.TabIndex = 6;
            // 
            // button_Select_TS_RecordMaster
            // 
            this.button_Select_TS_RecordMaster.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.button_Select_TS_RecordMaster.Location = new System.Drawing.Point(17, 121);
            this.button_Select_TS_RecordMaster.Name = "button_Select_TS_RecordMaster";
            this.button_Select_TS_RecordMaster.Size = new System.Drawing.Size(144, 44);
            this.button_Select_TS_RecordMaster.TabIndex = 5;
            this.button_Select_TS_RecordMaster.Text = "Select TS RecordMaster";
            this.button_Select_TS_RecordMaster.UseVisualStyleBackColor = true;
            this.button_Select_TS_RecordMaster.Click += new System.EventHandler(this.button_Select_TS_RecordMaster_Click);
            // 
            // button_ts_Transfer_TS_Model
            // 
            this.button_ts_Transfer_TS_Model.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.button_ts_Transfer_TS_Model.Location = new System.Drawing.Point(437, 225);
            this.button_ts_Transfer_TS_Model.Name = "button_ts_Transfer_TS_Model";
            this.button_ts_Transfer_TS_Model.Size = new System.Drawing.Size(287, 44);
            this.button_ts_Transfer_TS_Model.TabIndex = 4;
            this.button_ts_Transfer_TS_Model.Text = "Transfer TimeSheet to Master - Model";
            this.button_ts_Transfer_TS_Model.UseVisualStyleBackColor = true;
            this.button_ts_Transfer_TS_Model.Click += new System.EventHandler(this.button_ts_Transfer_TS_Model_Click);
            // 
            // button_ts_Transfer_TS_Drawing
            // 
            this.button_ts_Transfer_TS_Drawing.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.button_ts_Transfer_TS_Drawing.Location = new System.Drawing.Point(98, 225);
            this.button_ts_Transfer_TS_Drawing.Name = "button_ts_Transfer_TS_Drawing";
            this.button_ts_Transfer_TS_Drawing.Size = new System.Drawing.Size(287, 44);
            this.button_ts_Transfer_TS_Drawing.TabIndex = 3;
            this.button_ts_Transfer_TS_Drawing.Text = "Transfer TimeSheet to Master - Drawing";
            this.button_ts_Transfer_TS_Drawing.UseVisualStyleBackColor = true;
            this.button_ts_Transfer_TS_Drawing.Click += new System.EventHandler(this.button_ts_Transfer_TS_Drawing_Click);
            // 
            // button_ts_Select_InputTS
            // 
            this.button_ts_Select_InputTS.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.button_ts_Select_InputTS.Location = new System.Drawing.Point(17, 71);
            this.button_ts_Select_InputTS.Name = "button_ts_Select_InputTS";
            this.button_ts_Select_InputTS.Size = new System.Drawing.Size(144, 44);
            this.button_ts_Select_InputTS.TabIndex = 2;
            this.button_ts_Select_InputTS.Text = "Select Input TimeSheet file";
            this.button_ts_Select_InputTS.UseVisualStyleBackColor = true;
            this.button_ts_Select_InputTS.Click += new System.EventHandler(this.button_ts_Select_InputTS_Click);
            // 
            // button_ts_Select_Master
            // 
            this.button_ts_Select_Master.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.button_ts_Select_Master.Location = new System.Drawing.Point(17, 21);
            this.button_ts_Select_Master.Name = "button_ts_Select_Master";
            this.button_ts_Select_Master.Size = new System.Drawing.Size(144, 44);
            this.button_ts_Select_Master.TabIndex = 1;
            this.button_ts_Select_Master.Text = "Select Master file";
            this.button_ts_Select_Master.UseVisualStyleBackColor = true;
            this.button_ts_Select_Master.Click += new System.EventHandler(this.button_ts_Select_Master_Click);
            // 
            // tabPage3
            // 
            this.tabPage3.Controls.Add(this.button_Output_Form);
            this.tabPage3.Location = new System.Drawing.Point(4, 29);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage3.Size = new System.Drawing.Size(887, 588);
            this.tabPage3.TabIndex = 2;
            this.tabPage3.Text = "Output Form";
            // 
            // button_Output_Form
            // 
            this.button_Output_Form.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.button_Output_Form.Location = new System.Drawing.Point(8, 57);
            this.button_Output_Form.Name = "button_Output_Form";
            this.button_Output_Form.Size = new System.Drawing.Size(287, 44);
            this.button_Output_Form.TabIndex = 4;
            this.button_Output_Form.Text = "Transfer data from InputData to Output form";
            this.button_Output_Form.UseVisualStyleBackColor = true;
            this.button_Output_Form.Click += new System.EventHandler(this.button_Output_Form_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.Red;
            this.label1.Location = new System.Drawing.Point(262, 4);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(587, 17);
            this.label1.TabIndex = 20;
            this.label1.Text = "Please note to only transfer data with the \'Update Status\' column marked as OK";
            // 
            // FormMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(895, 621);
            this.Controls.Add(this.tabControl1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FormMain";
            this.Text = "Transfer data from Input to Master file";
            this.TopMost = true;
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tableLayoutPanel2.ResumeLayout(false);
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            this.tabPage3.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button button_Select_Master;
        private System.Windows.Forms.Button button_Check_InputData_Drawing;
        private System.Windows.Forms.TextBox txtFilePath_MasterFile;
        private System.Windows.Forms.Button button_Select_InputData;
        private System.Windows.Forms.TextBox txtFilePath_InputDataFile;
        private System.Windows.Forms.Button button_Select_InputRecordMaster;
        private System.Windows.Forms.TextBox txtFilePath_InputRecordMasterFile;
        private System.Windows.Forms.Button button_Transfer_RLMU_Receive_1st;
        private System.Windows.Forms.Button button_Transfer_RLMU_Receive_Final;
        private System.Windows.Forms.Button button_Transfer_Drawing_Submit_1st;
        private System.Windows.Forms.Button button_Transfer_Drawing_Submit_Final;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.ComponentModel.BackgroundWorker backgroundWorker1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Button button_Transfer_Model_Receive;
        private System.Windows.Forms.Button button_Transfer_Model_Submit;
        private System.Windows.Forms.Label labelStatus;
        private System.Windows.Forms.Button button_Transfer_Drawing_Feedback_Final;
        private System.Windows.Forms.Button button_Transfer_Drawing_Feedback_1st;
        private System.Windows.Forms.Button button_Transfer_Model_Feedback;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel2;
        private System.Windows.Forms.Button button_ts_Select_Master;
        private System.Windows.Forms.Button button_Check_InputData_Model;
        private System.Windows.Forms.Button button_ts_Transfer_TS_Drawing;
        private System.Windows.Forms.Button button_ts_Select_InputTS;
        private System.Windows.Forms.Button button_ts_Transfer_TS_Model;
        private System.Windows.Forms.Button button_Select_TS_RecordMaster;
        private System.Windows.Forms.TextBox txtFilePath_TS_MasterFile;
        private System.Windows.Forms.TextBox txtFilePath_TS_InputRecordMasterFile;
        private System.Windows.Forms.TextBox txtFilePath_TS_InputDataFile;
        private System.Windows.Forms.Label labelStatus_TS;
        private System.Windows.Forms.Button button_Transfer_Drawing_RFI_1st;
        private System.Windows.Forms.Button button_Transfer_Drawing_RFI_Final;
        private System.Windows.Forms.TabPage tabPage3;
        private System.Windows.Forms.Button button_Output_Form;
        private System.Windows.Forms.Button button_Transfer_Working_Receive;
        private System.Windows.Forms.Label label1;
    }
}