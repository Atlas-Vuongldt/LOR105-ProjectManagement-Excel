namespace ExcelTracking
{
    partial class GetInputData
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(GetInputData));
            this.button_GetFileNameFromInputFolder = new System.Windows.Forms.Button();
            this.IsInputAsModel = new System.Windows.Forms.CheckBox();
            this.IsInputAsPTANo = new System.Windows.Forms.CheckBox();
            this.button_GetFile_Select_Master = new System.Windows.Forms.Button();
            this.txtFilePath_GetFile_MasterFile = new System.Windows.Forms.TextBox();
            this.button_GetFile_Select_InputFolder = new System.Windows.Forms.Button();
            this.txtFilePath_GetFile_InputDataFolder = new System.Windows.Forms.TextBox();
            this.button_GetFile_Select_OutputFolder = new System.Windows.Forms.Button();
            this.txtFilePath_GetFile_OutputDataFolder = new System.Windows.Forms.TextBox();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.label2 = new System.Windows.Forms.Label();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.label1 = new System.Windows.Forms.Label();
            this.button_Output_Form = new System.Windows.Forms.Button();
            this.Open_CadConfigFolder = new System.Windows.Forms.Button();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.SuspendLayout();
            // 
            // button_GetFileNameFromInputFolder
            // 
            this.button_GetFileNameFromInputFolder.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.button_GetFileNameFromInputFolder.Location = new System.Drawing.Point(278, 254);
            this.button_GetFileNameFromInputFolder.Name = "button_GetFileNameFromInputFolder";
            this.button_GetFileNameFromInputFolder.Size = new System.Drawing.Size(287, 44);
            this.button_GetFileNameFromInputFolder.TabIndex = 12;
            this.button_GetFileNameFromInputFolder.Text = "Get FileName from Input Folder";
            this.button_GetFileNameFromInputFolder.UseVisualStyleBackColor = true;
            this.button_GetFileNameFromInputFolder.Click += new System.EventHandler(this.button_GetFileNameFromInputFolder_Click);
            // 
            // IsInputAsModel
            // 
            this.IsInputAsModel.AutoSize = true;
            this.IsInputAsModel.Location = new System.Drawing.Point(322, 216);
            this.IsInputAsModel.Name = "IsInputAsModel";
            this.IsInputAsModel.Size = new System.Drawing.Size(195, 17);
            this.IsInputAsModel.TabIndex = 24;
            this.IsInputAsModel.Text = "Input as Model (Default as Drawing)";
            this.IsInputAsModel.UseVisualStyleBackColor = true;
            // 
            // IsInputAsPTANo
            // 
            this.IsInputAsPTANo.AutoSize = true;
            this.IsInputAsPTANo.Location = new System.Drawing.Point(322, 195);
            this.IsInputAsPTANo.Name = "IsInputAsPTANo";
            this.IsInputAsPTANo.Size = new System.Drawing.Size(208, 17);
            this.IsInputAsPTANo.TabIndex = 23;
            this.IsInputAsPTANo.Text = "Input as PTA No. (Default as Doc Ref)";
            this.IsInputAsPTANo.UseVisualStyleBackColor = true;
            // 
            // button_GetFile_Select_Master
            // 
            this.button_GetFile_Select_Master.Location = new System.Drawing.Point(1, 129);
            this.button_GetFile_Select_Master.Name = "button_GetFile_Select_Master";
            this.button_GetFile_Select_Master.Size = new System.Drawing.Size(144, 44);
            this.button_GetFile_Select_Master.TabIndex = 21;
            this.button_GetFile_Select_Master.Text = "Select Master file";
            this.button_GetFile_Select_Master.UseVisualStyleBackColor = true;
            this.button_GetFile_Select_Master.Click += new System.EventHandler(this.button_GetFile_Select_Master_Click);
            // 
            // txtFilePath_GetFile_MasterFile
            // 
            this.txtFilePath_GetFile_MasterFile.BackColor = System.Drawing.Color.White;
            this.txtFilePath_GetFile_MasterFile.ForeColor = System.Drawing.Color.Red;
            this.txtFilePath_GetFile_MasterFile.Location = new System.Drawing.Point(160, 129);
            this.txtFilePath_GetFile_MasterFile.Multiline = true;
            this.txtFilePath_GetFile_MasterFile.Name = "txtFilePath_GetFile_MasterFile";
            this.txtFilePath_GetFile_MasterFile.ReadOnly = true;
            this.txtFilePath_GetFile_MasterFile.Size = new System.Drawing.Size(725, 44);
            this.txtFilePath_GetFile_MasterFile.TabIndex = 22;
            // 
            // button_GetFile_Select_InputFolder
            // 
            this.button_GetFile_Select_InputFolder.Location = new System.Drawing.Point(1, 17);
            this.button_GetFile_Select_InputFolder.Name = "button_GetFile_Select_InputFolder";
            this.button_GetFile_Select_InputFolder.Size = new System.Drawing.Size(144, 44);
            this.button_GetFile_Select_InputFolder.TabIndex = 19;
            this.button_GetFile_Select_InputFolder.Text = "Select Input Folder";
            this.button_GetFile_Select_InputFolder.UseVisualStyleBackColor = true;
            this.button_GetFile_Select_InputFolder.Click += new System.EventHandler(this.button_GetFile_Select_InputFolder_Click);
            // 
            // txtFilePath_GetFile_InputDataFolder
            // 
            this.txtFilePath_GetFile_InputDataFolder.BackColor = System.Drawing.Color.White;
            this.txtFilePath_GetFile_InputDataFolder.ForeColor = System.Drawing.Color.Red;
            this.txtFilePath_GetFile_InputDataFolder.Location = new System.Drawing.Point(160, 17);
            this.txtFilePath_GetFile_InputDataFolder.Multiline = true;
            this.txtFilePath_GetFile_InputDataFolder.Name = "txtFilePath_GetFile_InputDataFolder";
            this.txtFilePath_GetFile_InputDataFolder.ReadOnly = true;
            this.txtFilePath_GetFile_InputDataFolder.Size = new System.Drawing.Size(725, 44);
            this.txtFilePath_GetFile_InputDataFolder.TabIndex = 20;
            // 
            // button_GetFile_Select_OutputFolder
            // 
            this.button_GetFile_Select_OutputFolder.Location = new System.Drawing.Point(1, 71);
            this.button_GetFile_Select_OutputFolder.Name = "button_GetFile_Select_OutputFolder";
            this.button_GetFile_Select_OutputFolder.Size = new System.Drawing.Size(144, 44);
            this.button_GetFile_Select_OutputFolder.TabIndex = 17;
            this.button_GetFile_Select_OutputFolder.Text = "Select Output Folder";
            this.button_GetFile_Select_OutputFolder.UseVisualStyleBackColor = true;
            this.button_GetFile_Select_OutputFolder.Click += new System.EventHandler(this.button_GetFile_Select_OutputFolder_Click);
            // 
            // txtFilePath_GetFile_OutputDataFolder
            // 
            this.txtFilePath_GetFile_OutputDataFolder.BackColor = System.Drawing.Color.White;
            this.txtFilePath_GetFile_OutputDataFolder.ForeColor = System.Drawing.Color.Red;
            this.txtFilePath_GetFile_OutputDataFolder.Location = new System.Drawing.Point(160, 71);
            this.txtFilePath_GetFile_OutputDataFolder.Multiline = true;
            this.txtFilePath_GetFile_OutputDataFolder.Name = "txtFilePath_GetFile_OutputDataFolder";
            this.txtFilePath_GetFile_OutputDataFolder.ReadOnly = true;
            this.txtFilePath_GetFile_OutputDataFolder.Size = new System.Drawing.Size(725, 44);
            this.txtFilePath_GetFile_OutputDataFolder.TabIndex = 18;
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.ItemSize = new System.Drawing.Size(80, 25);
            this.tabControl1.Location = new System.Drawing.Point(3, 2);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(893, 391);
            this.tabControl1.TabIndex = 25;
            // 
            // tabPage1
            // 
            this.tabPage1.BackColor = System.Drawing.Color.Bisque;
            this.tabPage1.Controls.Add(this.Open_CadConfigFolder);
            this.tabPage1.Controls.Add(this.label2);
            this.tabPage1.Controls.Add(this.IsInputAsPTANo);
            this.tabPage1.Controls.Add(this.button_GetFile_Select_Master);
            this.tabPage1.Controls.Add(this.IsInputAsModel);
            this.tabPage1.Controls.Add(this.txtFilePath_GetFile_MasterFile);
            this.tabPage1.Controls.Add(this.button_GetFileNameFromInputFolder);
            this.tabPage1.Controls.Add(this.button_GetFile_Select_InputFolder);
            this.tabPage1.Controls.Add(this.txtFilePath_GetFile_InputDataFolder);
            this.tabPage1.Controls.Add(this.txtFilePath_GetFile_OutputDataFolder);
            this.tabPage1.Controls.Add(this.button_GetFile_Select_OutputFolder);
            this.tabPage1.Location = new System.Drawing.Point(4, 29);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(885, 358);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Get Info";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.Color.Red;
            this.label2.Location = new System.Drawing.Point(94, 311);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(700, 16);
            this.label2.TabIndex = 25;
            this.label2.Text = "The data obtained from the files inside the Input folder will be inserted into Sh" +
    "eet 2 (GetFilesInFolder)";
            // 
            // tabPage2
            // 
            this.tabPage2.BackColor = System.Drawing.Color.Bisque;
            this.tabPage2.Controls.Add(this.label1);
            this.tabPage2.Controls.Add(this.button_Output_Form);
            this.tabPage2.Location = new System.Drawing.Point(4, 29);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(885, 358);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Output Form";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.Red;
            this.label1.Location = new System.Drawing.Point(168, 99);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(577, 16);
            this.label1.TabIndex = 7;
            this.label1.Text = "Data will be transferred from Sheet 1 (Template Input Data) to Sheet 3 (OutputFro" +
    "m)";
            // 
            // button_Output_Form
            // 
            this.button_Output_Form.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.button_Output_Form.Location = new System.Drawing.Point(307, 37);
            this.button_Output_Form.Name = "button_Output_Form";
            this.button_Output_Form.Size = new System.Drawing.Size(287, 44);
            this.button_Output_Form.TabIndex = 6;
            this.button_Output_Form.Text = "Transfer data from InputData to Output form";
            this.button_Output_Form.UseVisualStyleBackColor = true;
            this.button_Output_Form.Click += new System.EventHandler(this.button_Output_Form_Click);
            // 
            // Open_CadConfigFolder
            // 
            this.Open_CadConfigFolder.Location = new System.Drawing.Point(1, 189);
            this.Open_CadConfigFolder.Name = "Open_CadConfigFolder";
            this.Open_CadConfigFolder.Size = new System.Drawing.Size(157, 23);
            this.Open_CadConfigFolder.TabIndex = 26;
            this.Open_CadConfigFolder.Text = "Open CadBlockConfig Folder";
            this.Open_CadConfigFolder.UseVisualStyleBackColor = true;
            this.Open_CadConfigFolder.Click += new System.EventHandler(this.Open_CadConfigFolder_Click);
            // 
            // GetInputData
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(903, 393);
            this.Controls.Add(this.tabControl1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "GetInputData";
            this.Text = "Get Files & Data in Folder";
            this.TopMost = true;
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button button_GetFileNameFromInputFolder;
        private System.Windows.Forms.CheckBox IsInputAsModel;
        private System.Windows.Forms.CheckBox IsInputAsPTANo;
        private System.Windows.Forms.Button button_GetFile_Select_Master;
        private System.Windows.Forms.TextBox txtFilePath_GetFile_MasterFile;
        private System.Windows.Forms.Button button_GetFile_Select_InputFolder;
        private System.Windows.Forms.TextBox txtFilePath_GetFile_InputDataFolder;
        private System.Windows.Forms.Button button_GetFile_Select_OutputFolder;
        private System.Windows.Forms.TextBox txtFilePath_GetFile_OutputDataFolder;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.Button button_Output_Form;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button Open_CadConfigFolder;
    }
}