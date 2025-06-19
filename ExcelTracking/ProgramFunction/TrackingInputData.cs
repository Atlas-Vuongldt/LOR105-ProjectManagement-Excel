#region using
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Forms;
using ExcelTracking;
using ExcelDataManager;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using System.Runtime.CompilerServices;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Window;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.Style;
using Autodesk.AutoCAD.DatabaseServices;
using static ExcelTracking.MainFunction;
using Application = System.Windows.Forms.Application;
using MessageBox = System.Windows.Forms.MessageBox;
#endregion

// NOTE
// Bổ sung phần đưa RedlineMarkups và PackageStamp vào Master và backup vào Master Record

namespace ExcelTracking
{
    public class TrackingInputData
    {
        public static int InputData_WorksheetIndex_Main = 1;
        public static Form activeForm = Application.OpenForms["FormMain"];

        public static Form GetActiveForm()
        {
            if (Application.OpenForms["FormMain"] != null)
            {
                var form = Application.OpenForms["FormMain"];

                // Kiểm tra nếu đang ở thread khác
                if (form.InvokeRequired)
                {
                    return (Form)form.Invoke(new Func<Form>(() => GetActiveForm()));
                }

                return form;
            }

            return Application.OpenForms.Cast<Form>().FirstOrDefault() ??
                   new Form() { TopMost = true };
        }

        // Main Tab
        public static string txtFilePath_Master = "";
        public static string txtFilePath_InputData = "";
        public static string txtFilePath_InputRecordMaster = "";

        // TimeSheet Tab
        public static string txtFilePath_TS_InputData = "";
        public static string txtFilePath_TS_InputRecordMaster = "";

        // GetFilesInFolder Tab
        public static string txtFilePath_GetFile_InputDataFolder = "";
        public static string txtFilePath_GetFile_OutputDataFolder = "";
        public static bool isInputAsModel = false;
        public static bool isInputAsPTANo = false;
        public static bool isGetCADInfo = false;
        // Others
        public static string status_DuplicateDocRef = "Duplicate DocRef";

        public static string InputDataTemplate_Sheet_Drawing = "Template Input Data";
        //============================================================================
        // ======CHECKING DATA
        public static void CheckInputDataFile_Drawing()
        {
            #region
            if (string.IsNullOrEmpty(txtFilePath_Master) || string.IsNullOrEmpty(txtFilePath_InputData) || string.IsNullOrEmpty(txtFilePath_InputRecordMaster))
            { 
                MessageBox.Show(activeForm, "Please select all the required Excel files!", "Warning", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return; 
            }

            // ========================== Process
            FileInfo fileInfo = new FileInfo(txtFilePath_InputData);
            if (!fileInfo.Exists)
            {
                MessageBox.Show(activeForm, "The Excel file" + Path.GetFileName(txtFilePath_InputData) + " does not exist!", "Warning", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            //--------------------------------------------
            using (var packageMaster = new ExcelPackage(new FileInfo(txtFilePath_Master)))
            using (var packageInput = new ExcelPackage(new FileInfo(txtFilePath_InputData)))
            using (var packageInputRecordMaster = new ExcelPackage(new FileInfo(txtFilePath_InputRecordMaster)))
            {
                // * Ktra có Sheet cần tìm trong Master file ko rồi mới bắt đầu
                var wsMaster = GetWorksheetByName(packageMaster, MasterExcelData_Drawing.SheetName);
                if (wsMaster == null)
                {
                    MessageBox.Show(activeForm, "Không tìm thấy sheet " + MasterExcelData_Drawing.SheetName + " trong Master file", "Error", 
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // * Thông tin các file excel
                var wsInput = packageInput.Workbook.Worksheets[InputData_WorksheetIndex_Main];
                var wsInputRecordMaster = packageInputRecordMaster.Workbook.Worksheets[1];

                int totalRowsInput = FindLastRowWithDataInColumn(wsInput, 2, InputExcelData.Start_Row);
                int totalRowsMaster = wsMaster.Dimension.End.Row;
                int totalRowsInputRecordMaster = wsInputRecordMaster.Dimension.End.Row;

                // --------------------------------------------------------------
                // *** Phần 1: TRANS ID / Date Receive vào cột tương ứng
                Copy_TransID_Date(wsInput, totalRowsInput);

                // --------------------------------------------------------------
                // *** Phần 2: Kiểm tra trùng lặp Doc Ref trong file InputData
                Check_Duplicate_DocRef_InputData(wsInput, totalRowsInput);

                // --------------------------------------------------------------
                // *** Phần 3: Đếm số lần xuất hiện của giá trị DocRef từ InputRecordMaster và điền số lần tiếp theo vào Input Data
                TotalTimesCount_DocRef_InputRecordMaster(
                    wsInput, InputExcelData.DocRef_Col, InputExcelData.TimesCount_Col, InputExcelData.Start_Row, totalRowsInput, 
                    wsInputRecordMaster, totalRowsInputRecordMaster);

                // --------------------------------------------------------------
                // *** Phần 4: Kiểm tra TRANS ID đã xuất hiện trong InputRecordMaster chưa
                Check_Duplicate_TransID_InputRecordMaster(wsInput, totalRowsInput, wsInputRecordMaster, totalRowsInputRecordMaster);

                // --------------------------------------------------------------
                // *** Phần 5: Kiểm tra Doc Title trong Input file và trong Master file,
                // nếu giống thì OK, nếu khác thì note vào cột DocTitle trong Input file là "Updated to Master"
                // - Lấy giá trị Package và Discipline đưa vào Input Data
                for (int row = InputExcelData.Start_Row; row <= totalRowsInput; row++)
                {
                    string docRef = wsInput.Cells[row, InputExcelData.DocRef_Col].Text;             // Lấy giá trị Doc Ref từ Input file
                    string docTitleInput = wsInput.Cells[row, InputExcelData.DocTitle_Col].Text;    // Lấy giá trị Doc Title từ Input file

                    if (!string.IsNullOrEmpty(docRef) && !string.IsNullOrEmpty(docTitleInput))
                    {
                        for (int masterRow = MasterExcelData_Drawing.Start_Row; masterRow <= totalRowsMaster; masterRow++)
                        {
                            string allianceNo = wsMaster.Cells[masterRow, MasterExcelData_Drawing.Alliance_Col].Text;       // Lấy giá trị Alliance No trong Master
                            string docTitleMaster = wsMaster.Cells[masterRow, MasterExcelData_Drawing.DocTitle_Col].Text;   // Lấy giá trị Doc Title từ Master file

                            if (allianceNo == docRef) // Nếu Doc Ref trùng với Alliance No
                            {
                                // * Kiểm tra trùng lặp
                                if (!docTitleInput.Equals(docTitleMaster, StringComparison.OrdinalIgnoreCase))
                                {
                                    wsInput.Cells[row, InputExcelData.CheckDocTitle_Col].Value = "Need to update to Master";
                                    wsInput.Cells[row, InputExcelData.CheckDocTitle_Col].Style.Font.Color.SetColor(Color.Blue);
                                }
                                else
                                {
                                    wsInput.Cells[row, InputExcelData.CheckDocTitle_Col].Value = "OK";
                                }

                                // * Lấy giá trị Package và Discipline đưa vào Input Data
                                string packageValue = wsMaster.Cells[masterRow, MasterExcelData_Drawing.Package_Col].Text;          // Cột Package trong Master
                                string disciplineValue = wsMaster.Cells[masterRow, MasterExcelData_Drawing.Discipline_Col].Text;    // Cột Discipline trong Master

                                // Ghi vào Input Data
                                wsInput.Cells[row, InputExcelData.Package_Col].Value = packageValue;
                                wsInput.Cells[row, InputExcelData.Discipline_Col].Value = disciplineValue;

                                break;
                            }
                        }
                    }
                }

                // --------------------------------------------------------------
                // *** Phần 6: Xác định Doc Ref có trong Master file ko
                for (int row = InputExcelData.Start_Row; row <= totalRowsInput; row++)
                {
                    if (wsInput.Cells[row, InputExcelData.UpdateStatus_Col].Text == status_DuplicateDocRef)
                    {
                        continue; // Bỏ qua dòng Duplicate
                    }

                    string docRef = wsInput.Cells[row, InputExcelData.DocRef_Col].Text;
                    bool found = false;

                    for (int masterRow = MasterExcelData_Drawing.Start_Row; masterRow <= totalRowsMaster; masterRow++)
                    {
                        string allianceNo = wsMaster.Cells[masterRow, MasterExcelData_Drawing.Alliance_Col].Text;
                        if (allianceNo == docRef)
                        {
                            found = true;
                            break;
                        }
                    }

                    if (!found)
                    {
                        wsInput.Cells[row, InputExcelData.UpdateStatus_Col].Value = "Not Found DocRef in the Matster file";
                        wsInput.Cells[row, InputExcelData.UpdateStatus_Col].Style.Font.Color.SetColor(Color.Red);
                    }
                }

                // --------------------------------------------------------------
                // Phần cuối: Lưu các file
                // Lấy timestamp hiện tại
                string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                // * Tạo file _Checked cho Input Data, giữ nguyên tệp gốc
                string directory = Path.GetDirectoryName(txtFilePath_InputData);                        // Thư mục chứa tệp gốc
                string fileNameWithoutExt = Path.GetFileNameWithoutExtension(txtFilePath_InputData);    // Tên file không có đuôi
                string fileExtension = Path.GetExtension(txtFilePath_InputData);                        // Lấy đuôi file (.xlsx)
                string outputFilePath = Path.Combine(directory, $"{fileNameWithoutExt}_Checked_{timestamp}{fileExtension}");

                packageInput.SaveAs(new FileInfo(outputFilePath));

                MessageBox.Show(activeForm, "Data has been checked!", "Check Input Data", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            #endregion
        }
        public static void CheckInputDataFile_Model()
        {
            #region
            if (string.IsNullOrEmpty(txtFilePath_Master) || string.IsNullOrEmpty(txtFilePath_InputData) || string.IsNullOrEmpty(txtFilePath_InputRecordMaster))
            {
                MessageBox.Show(activeForm, "Please select all the required Excel files!", "Warning",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // ========================== Process
            FileInfo fileInfo = new FileInfo(txtFilePath_InputData);
            if (!fileInfo.Exists)
            {
                MessageBox.Show(activeForm, "The Excel file" + Path.GetFileName(txtFilePath_InputData) + " does not exist!", "Warning",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            //--------------------------------------------
            using (var packageMaster = new ExcelPackage(new FileInfo(txtFilePath_Master)))
            using (var packageInput = new ExcelPackage(new FileInfo(txtFilePath_InputData)))
            using (var packageInputRecordMaster = new ExcelPackage(new FileInfo(txtFilePath_InputRecordMaster)))
            {
                // * Ktra có Sheet cần tìm trong Master file ko rồi mới bắt đầu
                var wsMaster = GetWorksheetByName(packageMaster, MasterExcelData_Model.SheetName);
                if (wsMaster == null)
                {
                    MessageBox.Show(activeForm, "Không tìm thấy sheet " + MasterExcelData_Model.SheetName + " trong Master file", "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // * Thông tin các file excel
                var wsInput = packageInput.Workbook.Worksheets[InputData_WorksheetIndex_Main];
                var wsInputRecordMaster = packageInputRecordMaster.Workbook.Worksheets[1];

                int totalRowsInput = FindLastRowWithDataInColumn(wsInput, 2, InputExcelData.Start_Row);
                int totalRowsMaster = wsMaster.Dimension.End.Row;
                int totalRowsInputRecordMaster = wsInputRecordMaster.Dimension.End.Row;

                // --------------------------------------------------------------
                // *** Phần 1: TRANS ID / Date Receive vào cột tương ứng
                Copy_TransID_Date(wsInput, totalRowsInput);

                // --------------------------------------------------------------
                // *** Phần 2: Kiểm tra trùng lặp Doc Ref trong file InputData
                Check_Duplicate_DocRef_InputData(wsInput, totalRowsInput);

                // --------------------------------------------------------------
                // *** Phần 3: Đếm số lần xuất hiện của giá trị DocRef từ InputRecordMaster và điền số lần tiếp theo vào Input Data
                TotalTimesCount_DocRef_InputRecordMaster(
                    wsInput, InputExcelData.DocRef_Col, InputExcelData.TimesCount_Col, InputExcelData.Start_Row, totalRowsInput, 
                    wsInputRecordMaster, totalRowsInputRecordMaster);

                // --------------------------------------------------------------
                // *** Phần 4: Kiểm tra TRANS ID đã xuất hiện trong InputRecordMaster chưa
                Check_Duplicate_TransID_InputRecordMaster(wsInput, totalRowsInput, wsInputRecordMaster, totalRowsInputRecordMaster);

                // --------------------------------------------------------------
                // *** Phần 5: Kiểm tra Doc Title trong Input file và trong Master file,
                // nếu giống thì OK, nếu khác thì note vào cột DocTitle trong Input file là "Updated to Master"
                // - Lấy giá trị Package và Discipline đưa vào Input Data
                for (int row = InputExcelData.Start_Row; row <= totalRowsInput; row++)
                {
                    string docRef = wsInput.Cells[row, InputExcelData.DocRef_Col].Text;             // Lấy giá trị Doc Ref từ Input file
                    string docTitleInput = wsInput.Cells[row, InputExcelData.DocTitle_Col].Text;    // Lấy giá trị Doc Title từ Input file

                    if (!string.IsNullOrEmpty(docRef) && !string.IsNullOrEmpty(docTitleInput))
                    {
                        for (int masterRow = MasterExcelData_Model.Start_Row; masterRow <= totalRowsMaster; masterRow++)
                        {
                            string allianceNo = wsMaster.Cells[masterRow, MasterExcelData_Model.Alliance_Col].Text;       // Lấy giá trị Alliance No trong Master
                            string docTitleMaster = wsMaster.Cells[masterRow, MasterExcelData_Model.DocTitle_Col].Text;   // Lấy giá trị Doc Title từ Master file

                            if (allianceNo == docRef) // Nếu Doc Ref trùng với Alliance No
                            {
                                // * Kiểm tra trùng lặp
                                if (!docTitleInput.Equals(docTitleMaster, StringComparison.OrdinalIgnoreCase))
                                {
                                    wsInput.Cells[row, InputExcelData.CheckDocTitle_Col].Value = "Need to update to Master";
                                    wsInput.Cells[row, InputExcelData.CheckDocTitle_Col].Style.Font.Color.SetColor(Color.Blue);
                                }
                                else
                                {
                                    wsInput.Cells[row, InputExcelData.CheckDocTitle_Col].Value = "OK";
                                }

                                // * Lấy giá trị Package và Discipline đưa vào Input Data
                                string packageValue = wsMaster.Cells[masterRow, MasterExcelData_Model.Package_Col].Text;          // Cột Package trong Master
                                string disciplineValue = wsMaster.Cells[masterRow, MasterExcelData_Model.Discipline_Col].Text;    // Cột Discipline trong Master

                                // Ghi vào Input Data
                                wsInput.Cells[row, InputExcelData.Package_Col].Value = packageValue;
                                wsInput.Cells[row, InputExcelData.Discipline_Col].Value = disciplineValue;

                                break;
                            }
                        }
                    }
                }

                // --------------------------------------------------------------
                // *** Phần 6: Xác định Doc Ref có trong Master file ko
                for (int row = InputExcelData.Start_Row; row <= totalRowsInput; row++)
                {
                    if (wsInput.Cells[row, InputExcelData.UpdateStatus_Col].Text == status_DuplicateDocRef)
                    {
                        continue; // Bỏ qua dòng Duplicate
                    }

                    string docRef = wsInput.Cells[row, InputExcelData.DocRef_Col].Text;
                    bool found = false;

                    for (int masterRow = MasterExcelData_Model.Start_Row; masterRow <= totalRowsMaster; masterRow++)
                    {
                        string allianceNo = wsMaster.Cells[masterRow, MasterExcelData_Model.Alliance_Col].Text;
                        if (allianceNo == docRef)
                        {
                            found = true;
                            break;
                        }
                    }

                    if (!found)
                    {
                        wsInput.Cells[row, InputExcelData.UpdateStatus_Col].Value = "Not Found DocRef in the Matster file";
                        wsInput.Cells[row, InputExcelData.UpdateStatus_Col].Style.Font.Color.SetColor(Color.Red);
                    }
                }

                // --------------------------------------------------------------
                // Phần cuối: Lưu các file
                // Lấy timestamp hiện tại
                string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                // * Tạo file _Checked cho Input Data, giữ nguyên tệp gốc
                string directory = Path.GetDirectoryName(txtFilePath_InputData);                        // Thư mục chứa tệp gốc
                string fileNameWithoutExt = Path.GetFileNameWithoutExtension(txtFilePath_InputData);    // Tên file không có đuôi
                string fileExtension = Path.GetExtension(txtFilePath_InputData);                        // Lấy đuôi file (.xlsx)
                string outputFilePath = Path.Combine(directory, $"{fileNameWithoutExt}_Checked_{timestamp}{fileExtension}");

                packageInput.SaveAs(new FileInfo(outputFilePath));

                MessageBox.Show(activeForm, "Data has been checked!", "Check Input Data", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            #endregion
        }

        //-------------------------------------------------------
        public static void Copy_TransID_Date(ExcelWorksheet wsInput, int totalRowsInput)
        {
            #region
            string transId_Source = wsInput.Cells[InputExcelData.TRANSID_Source_Row, InputExcelData.TRANSID_Source_Col].Text;   // Lấy giá trị TRANS ID - Input file
            string dateReceiveValue_Source = wsInput.Cells[InputExcelData.DateReceive_Source_Row, InputExcelData.DateReceive_Source_Col].Text;  // Lấy giá trị Date - Input file

            for (int row = InputExcelData.Start_Row; row <= totalRowsInput; row++)
            {
                string docRef = wsInput.Cells[row, InputExcelData.DocRef_Col].Text;
                if (!string.IsNullOrEmpty(docRef))
                {
                    // * Gán TRANS ID vào cột TRANS ID cho tất cả DocRef trong InputData
                    wsInput.Cells[row, InputExcelData.TRANSID_Col].Value = transId_Source;
                    // * Gán Date Receive vào cột Date Receive cho tất cả DocRef trong InputData
                    DateTime dateValue;
                    if (DateTime.TryParseExact(dateReceiveValue_Source, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out dateValue))
                    {
                        wsInput.Cells[row, InputExcelData.Date_Col].Value = dateValue;
                        wsInput.Cells[row, InputExcelData.Date_Col].Style.Numberformat.Format = "dd/MM/yyyy"; // Định dạng ô theo ngày tháng năm
                    }
                    else
                    {
                        wsInput.Cells[row, InputExcelData.Date_Col].Value = dateReceiveValue_Source; // Giữ nguyên nếu không parse được
                    }
                }
            }
            #endregion
        }
        public static void Check_Duplicate_DocRef_InputData(ExcelWorksheet wsInput, int totalRowsInput)
        {
            #region
            var docRefList = wsInput.Cells[InputExcelData.Start_Row, InputExcelData.DocRef_Col, totalRowsInput, InputExcelData.DocRef_Col]
                                      .Select(cell => cell.Text).ToList();
            for (int row = InputExcelData.Start_Row; row <= totalRowsInput; row++)
            {
                string docRef = wsInput.Cells[row, InputExcelData.DocRef_Col].Text;
                if (!string.IsNullOrEmpty(docRef))
                {
                    // * Kiểm tra trùng lặp Doc Ref
                    int count = docRefList.Count(x => x == docRef);
                    if (count > 1)
                    {
                        wsInput.Cells[row, InputExcelData.UpdateStatus_Col].Value = status_DuplicateDocRef;
                        wsInput.Cells[row, InputExcelData.UpdateStatus_Col].Style.Font.Color.SetColor(Color.Red);
                    }
                    else
                    {
                        wsInput.Cells[row, InputExcelData.UpdateStatus_Col].Value = "OK";
                    }
                }
            }
            #endregion
        }
        public static void TotalTimesCount_DocRef_InputRecordMaster(
            ExcelWorksheet wsInput, int input_DocRef_Col, int input_TimesCount_Col, int input_StartRow, int input_TotalRows,
            ExcelWorksheet wsInputRecordMaster, int inputRecordMaster_TotalRows)
        {
            #region
            // Dictionary lưu số lần xuất hiện của DocRef trong Input Master
            Dictionary<string, int> docRefCounts = new Dictionary<string, int>();

            // Bước 1: Quét toàn bộ InputRecordMaster trước, lưu số lần xuất hiện vào Dictionary
            for (int rowMaster = InputRecordMasterExcelData.Start_Row; rowMaster <= inputRecordMaster_TotalRows; rowMaster++)
            {
                string docRef_InputRecordMaster = wsInputRecordMaster.Cells[rowMaster, InputRecordMasterExcelData.DocRef_Col].Text;

                if (!string.IsNullOrEmpty(docRef_InputRecordMaster))
                {
                    if (docRefCounts.ContainsKey(docRef_InputRecordMaster))
                        docRefCounts[docRef_InputRecordMaster]++;
                    else
                        docRefCounts[docRef_InputRecordMaster] = 1;
                }
            }

            // Bước 2: Duyệt từng dòng trong Input Data, cập nhật số lần xuất hiện
            for (int rowInput = input_StartRow; rowInput <= input_TotalRows; rowInput++)
            {
                string docRef = wsInput.Cells[rowInput, input_DocRef_Col].Text;

                if (!string.IsNullOrEmpty(docRef))
                {
                    // Nếu đã có trong Dictionary, tăng count lên 1 (lần tiếp theo)
                    if (docRefCounts.ContainsKey(docRef))
                    {
                        docRefCounts[docRef]++;
                    }
                    else
                    {
                        docRefCounts[docRef] = 1; // Lần đầu tiên xuất hiện
                    }

                    // Ghi giá trị Count vào cột Times Count trong Input Data
                    wsInput.Cells[rowInput, input_TimesCount_Col].Value = docRefCounts[docRef];
                }
            }
            #endregion
        }
        public static void Check_Duplicate_TransID_InputRecordMaster(ExcelWorksheet wsInput, int totalRowsInput,
            ExcelWorksheet wsInputRecordMaster, int totalRowsInputRecordMaster)
        {
            #region
            HashSet<string> transIdSet = new HashSet<string>();

            // Lưu danh sách TransID từ InputRecordMaster vào HashSet để kiểm tra nhanh
            for (int rowMaster = InputRecordMasterExcelData.Start_Row; rowMaster <= totalRowsInputRecordMaster; rowMaster++)
            {
                string transID_InputRecordMaster = wsInputRecordMaster.Cells[rowMaster, InputRecordMasterExcelData.TRANSID_Col].Text;
                if (!string.IsNullOrEmpty(transID_InputRecordMaster))
                {
                    transIdSet.Add(transID_InputRecordMaster);
                }
            }

            // Kiểm tra từng Trans ID trong Input file
            for (int row = InputExcelData.Start_Row; row <= totalRowsInput; row++)
            {
                string transId = wsInput.Cells[row, InputExcelData.TRANSID_Col].Text;

                if (!string.IsNullOrEmpty(transId) && transIdSet.Contains(transId))
                {
                    wsInput.Cells[row, InputExcelData.UpdateStatus_Col].Value = "Duplicate TransID";
                    wsInput.Cells[row, InputExcelData.UpdateStatus_Col].Style.Font.Color.SetColor(Color.Red);
                }
            }
            #endregion
        }
        //============================================================================
        // ======TRANSFER DATA - MAIN TAB

        //-------------------------------------------------------
        // DRAWING - WORKING
        public static void Transfer_WORKING_Receive_ToMaster()
        {
            string caption = "Transfer WORKING_Receive to Master";
            string sheetName_InputRecorMaster = InputRecordMasterExcelData_WORKING_Receive.sheetName;

            if (!IsValidExcelFiles_ForTransfer_MainTab(activeForm, caption, sheetName_InputRecorMaster)) { return; }

            //--------------------------------------------
            // *** Backup Master và InputRecordMaster trước khi transfer data
            // * Backup Master File
            BackupFileToBackupFolder(txtFilePath_Master, Path.Combine("_backup", "backup_Master"));

            // * Backup InputRecordMaster File
            BackupFileToBackupFolder(txtFilePath_InputRecordMaster, Path.Combine("_backup", "backup_" + sheetName_InputRecorMaster));

            //--------------------------------------------
            // *** Transfer data
            bool isInput_Modified = false;

            // * Lấy giá trị ngày giờ trong filename của InputData file
            string dateTime_Checked = "";
            string fileName_InputData = Path.GetFileName(txtFilePath_InputData);
            string[] parts = fileName_InputData.Split('_');

            if (parts.Length >= 4)
            {
                dateTime_Checked = parts[parts.Length - 2] + "_" + parts[parts.Length - 1].Replace(".xlsx", "");
            }

            using (var packageMaster = new ExcelPackage(new FileInfo(txtFilePath_Master)))
            using (var packageInput = new ExcelPackage(new FileInfo(txtFilePath_InputData)))
            using (var packageInputRecordMaster = new ExcelPackage(new FileInfo(txtFilePath_InputRecordMaster)))
            {
                // * Ktra có Sheet cần tìm trong Master file ko rồi mới bắt đầu
                var wsMaster = GetWorksheetByName(packageMaster, MasterExcelData_Drawing.SheetName);
                if (wsMaster == null)
                {
                    MessageBox.Show(activeForm, "Sheet " + MasterExcelData_Drawing.SheetName + " not found in the file " + Path.GetFileName(txtFilePath_Master), "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                var wsInput = packageInput.Workbook.Worksheets[InputData_WorksheetIndex_Main];
                var wsInputRecordMaster = GetWorksheetByName(packageInputRecordMaster, sheetName_InputRecorMaster);

                // * Thông tin các cột của Master file
                int docTitle_Master_Col = MasterExcelData_Drawing.DocTitle_Col;
                int transid_Master_Col = MasterExcelData_Drawing_WorkingReceive.TRANSID_Col;
                int rev_Master_Col = MasterExcelData_Drawing_WorkingReceive.REV_Col;
                int date_Receive_Master_Col = MasterExcelData_Drawing_WorkingReceive.Date_Receive_Col;
                int purpose_Master_Col = MasterExcelData_Drawing_WorkingReceive.Purpose_Col;
                int date_Issue_Master_Col = MasterExcelData_Drawing_WorkingReceive.Date_Issue_Col;
                int status_Master_Col = MasterExcelData_Drawing_WorkingReceive.Status_Col;
                int atlasComment_Master_Col = MasterExcelData_Drawing_WorkingReceive.AtlasComment_Col;

                // * Thông tin các cột của InputRecordMaster file
                int docRef_InputRecordMaster_Col = InputRecordMasterExcelData.DocRef_Col;
                int docTitle_InputRecordMaster_Col = InputRecordMasterExcelData.DocTitle_Col;
                int transid_InputRecordMaster_Col = InputRecordMasterExcelData.TRANSID_Col;

                int ver_InputRecordMaster_Col = -1;
                int rev_InputRecordMaster_Col = InputRecordMasterExcelData_WORKING_Receive.REV_Col;
                int purpose_InputRecordMaster_Col = InputRecordMasterExcelData_WORKING_Receive.Purpose_Col;
                int status_InputRecordMaster_Col = InputRecordMasterExcelData_WORKING_Receive.Status_Col;
                int date_Receive_InputRecordMaster_Col = InputRecordMasterExcelData_WORKING_Receive.Date_Receive_Col;
                int timesCount_InputRecordMaster_Col = InputRecordMasterExcelData_WORKING_Receive.TimesCount_Col;
                int package_InputRecordMaster_Col = -1;
                int discipline_InputRecordMaster_Col = -1;
                int redlineMarkup_InputRecordMaster_Col = -1;
                int packageStampStatus_InputRecordMaster_Col = -1;
                int date_Issue_InputRecordMaster_Col = InputRecordMasterExcelData_WORKING_Receive.Date_Issue_Col;
                int modelName_InputRecordMaster_Col = -1;
                int nativeFileType_InputRecordMaster_Col = -1;
                int submittedFileType_InputRecordMaster_Col = -1;
                int atlasComment_InputRecordMaster_Col = InputRecordMasterExcelData_WORKING_Receive.AtlasComment_Col;
                int updateStatus_InputRecordMaster_Col = -1;
                int checkDocTitle_InputRecordMaster_Col = -1;

                // * Lấy về tổng số dòng của các file
                int totalRowsInput = wsInput.Dimension.End.Row;
                int totalRowsMaster = wsMaster.Dimension.End.Row;

                // *** Đưa data vào Master
                for (int row = InputExcelData.Start_Row; row <= totalRowsInput; row++)
                {
                    // Chỉ lấy dòng có STATUS = OK
                    if (wsInput.Cells[row, InputExcelData.UpdateStatus_Col].Text != "OK")
                    {
                        continue; // Bỏ qua dòng Duplicate
                    }

                    string docRef = wsInput.Cells[row, InputExcelData.DocRef_Col].Text;     // Lấy giá trị "Doc Ref" trong Input file
                    string docTitle = wsInput.Cells[row, InputExcelData.DocTitle_Col].Text;

                    string revValue = wsInput.Cells[row, InputExcelData.REV_Col].Text;
                    string purposeValue = wsInput.Cells[row, InputExcelData.Purpose_Col].Text;
                    string statusValue = wsInput.Cells[row, InputExcelData.Status_Col].Text;
                    string transIdValue = wsInput.Cells[row, InputExcelData.TRANSID_Col].Text;
                    string dateValue_str = wsInput.Cells[row, InputExcelData.Date_Col].Text;
                    string timesCountValue = wsInput.Cells[row, InputExcelData.TimesCount_Col].Text;
                    string packageValue = wsInput.Cells[row, InputExcelData.Package_Col].Text;
                    string disciplineValue = wsInput.Cells[row, InputExcelData.Discipline_Col].Text;
                    string redlineMarkupValue = wsInput.Cells[row, InputExcelData.RedlineMarkup_Col].Text;
                    string packageStampStatusValue = wsInput.Cells[row, InputExcelData.PackageStampStatus_Col].Text;
                    string dateIssueValue_str = wsInput.Cells[row, InputExcelData.DateIssue_Col].Text;
                    string nativeFileTypeValue = wsInput.Cells[row, InputExcelData.NativeFileType_Col].Text;
                    string submittedFileTypeValue = wsInput.Cells[row, InputExcelData.SubmittedFileType_Col].Text;
                    string modelNameValue = wsInput.Cells[row, InputExcelData.ModelName_Col].Text;
                    string atlasCommentValue = wsInput.Cells[row, InputExcelData.AtlasComment_Col].Text;
                    string updateStatusValue = wsInput.Cells[row, InputExcelData.UpdateStatus_Col].Text;
                    string checkDocTitleValue = wsInput.Cells[row, InputExcelData.CheckDocTitle_Col].Text;

                    if (string.IsNullOrEmpty(docRef)) { continue; }

                    for (int masterRow = MasterExcelData_Drawing.Start_Row; masterRow <= totalRowsMaster; masterRow++)
                    {
                        string allianceNo = wsMaster.Cells[masterRow, MasterExcelData_Drawing.Alliance_Col].Text;   // Giá trị "Alliance No." trong Master file
                        if (allianceNo == docRef)
                        {
                            // TRANS ID
                            wsMaster.Cells[masterRow, transid_Master_Col].Value = transIdValue;
                            // REV
                            wsMaster.Cells[masterRow, rev_Master_Col].Value = revValue;
                            if (int.TryParse(revValue, out int revValue_int))
                            {
                                wsMaster.Cells[masterRow, rev_Master_Col].Value = revValue_int;
                            }
                            // PURPOSE
                            wsMaster.Cells[masterRow, purpose_Master_Col].Value = purposeValue;
                            // DATE RECEIVE
                            DateTime dateValue_ddmmyyy;
                            if (DateTime.TryParseExact(dateValue_str, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out dateValue_ddmmyyy))
                            {
                                wsMaster.Cells[masterRow, date_Receive_Master_Col].Value = dateValue_ddmmyyy;
                                wsMaster.Cells[masterRow, date_Receive_Master_Col].Style.Numberformat.Format = "dd/MM/yyyy"; // Định dạng ô theo ngày tháng năm
                            }
                            else
                            {
                                wsMaster.Cells[masterRow, date_Receive_Master_Col].Value = dateValue_str; // Giữ nguyên nếu không parse được
                            }

                            // DATE ISSUE
                            if (DateTime.TryParseExact(dateIssueValue_str, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out dateValue_ddmmyyy))
                            {
                                wsMaster.Cells[masterRow, date_Issue_Master_Col].Value = dateValue_ddmmyyy;
                                wsMaster.Cells[masterRow, date_Issue_Master_Col].Style.Numberformat.Format = "dd/MM/yyyy"; // Định dạng ô theo ngày tháng năm
                            }
                            else
                            {
                                wsMaster.Cells[masterRow, date_Issue_Master_Col].Value = dateIssueValue_str; // Giữ nguyên nếu không parse được
                            }

                            // STATUS
                            wsMaster.Cells[masterRow, status_Master_Col].Value = statusValue;
                            // ATLAS COMMENT
                            wsMaster.Cells[masterRow, atlasComment_Master_Col].Value = atlasCommentValue;

                            break;
                        }
                    }
                }

                // *** Đưa data vào InputRecordMaster
                TransferToInputRecordMaster(wsInput, wsInputRecordMaster,
                    docRef_InputRecordMaster_Col,
                    docTitle_InputRecordMaster_Col,
                    ver_InputRecordMaster_Col,
                    rev_InputRecordMaster_Col,
                    purpose_InputRecordMaster_Col,
                    status_InputRecordMaster_Col,
                    transid_InputRecordMaster_Col,
                    date_Receive_InputRecordMaster_Col,
                    timesCount_InputRecordMaster_Col,
                    package_InputRecordMaster_Col,
                    discipline_InputRecordMaster_Col,
                    redlineMarkup_InputRecordMaster_Col,
                    packageStampStatus_InputRecordMaster_Col,
                    date_Issue_InputRecordMaster_Col,
                    modelName_InputRecordMaster_Col,
                    nativeFileType_InputRecordMaster_Col,
                    submittedFileType_InputRecordMaster_Col,
                    atlasComment_InputRecordMaster_Col,
                    updateStatus_InputRecordMaster_Col,
                    checkDocTitle_InputRecordMaster_Col,
                    dateTime_Checked,
                    out isInput_Modified);

                // *** Phần cuối: Lưu các file
                packageMaster.Save();
                packageInputRecordMaster.Save();

                MessageBox.Show(activeForm, "Data transfer is done!", caption, MessageBoxButtons.OK, MessageBoxIcon.Information);

            }


        }

        //-------------------------------------------------------
        // 01. DRAWING - FOR 1ST
        public static void Transfer_RLMU_Receive_1st_ToMaster()
        {
            #region
            string caption = "Transfer RLMU_Receive_1st to Master";
            string sheetName_InputRecorMaster = InputRecordMasterExcelData_RLMU_Receive_1st.sheetName;

            if (!IsValidExcelFiles_ForTransfer_MainTab(activeForm, caption, sheetName_InputRecorMaster)) { return; }

            //--------------------------------------------
            // *** Backup Master và InputRecordMaster trước khi transfer data
            // * Backup Master File
            BackupFileToBackupFolder(txtFilePath_Master, Path.Combine("_backup", "backup_Master"));

            // * Backup InputRecordMaster File
            BackupFileToBackupFolder(txtFilePath_InputRecordMaster, Path.Combine("_backup", "backup_" + sheetName_InputRecorMaster));

            //--------------------------------------------
            // *** Transfer data
            bool isInput_Modified = false;

            // * Lấy giá trị ngày giờ trong filename của InputData file
            string dateTime_Checked = "";
            string fileName_InputData = Path.GetFileName(txtFilePath_InputData);
            string[] parts = fileName_InputData.Split('_');

            if (parts.Length >= 4)
            {
                dateTime_Checked = parts[parts.Length - 2] + "_" + parts[parts.Length - 1].Replace(".xlsx", "");
            }

            using (var packageMaster = new ExcelPackage(new FileInfo(txtFilePath_Master)))
            using (var packageInput = new ExcelPackage(new FileInfo(txtFilePath_InputData)))
            using (var packageInputRecordMaster = new ExcelPackage(new FileInfo(txtFilePath_InputRecordMaster)))
            {
                // * Ktra có Sheet cần tìm trong Master file ko rồi mới bắt đầu
                var wsMaster = GetWorksheetByName(packageMaster, MasterExcelData_Drawing.SheetName);
                if (wsMaster == null)
                {
                    MessageBox.Show(activeForm, "Sheet " + MasterExcelData_Drawing.SheetName + " not found in the file " + Path.GetFileName(txtFilePath_Master), "Error", 
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                var wsInput = packageInput.Workbook.Worksheets[InputData_WorksheetIndex_Main];
                var wsInputRecordMaster = GetWorksheetByName(packageInputRecordMaster, sheetName_InputRecorMaster);

                // * Thông tin các cột của Master file
                int docTitle_Master_Col = MasterExcelData_Drawing.DocTitle_Col;
                int timesCount_Master_Col = MasterExcelData_Drawing_FirstReceive.TimesCount_Col;
                int transid_Master_Col = MasterExcelData_Drawing_FirstReceive.TRANSID_Col;
                int rev_Master_Col = MasterExcelData_Drawing_FirstReceive.REV_Col;
                int ver_Master_Col = MasterExcelData_Drawing_FirstReceive.Ver_Col;
                int purpose_Master_Col = MasterExcelData_Drawing_FirstReceive.Purpose_Col;
                int date_Master_Col = MasterExcelData_Drawing_FirstReceive.Date_Col;
                int status_Master_Col = MasterExcelData_Drawing_FirstReceive.Status_Col;
                int atlasComment_Master_Col = MasterExcelData_Drawing_FirstReceive.AtlasComment_Col;
                int redlineMarkup_Master_Col = MasterExcelData_Drawing_FirstReceive.RedlineMarkup_Col;
                int packageStampStatus_Master_Col = MasterExcelData_Drawing_FirstReceive.PackageStampStatus_Col;

                // * Thông tin các cột của InputRecordMaster file
                int docRef_InputRecordMaster_Col = InputRecordMasterExcelData.DocRef_Col;
                int docTitle_InputRecordMaster_Col = InputRecordMasterExcelData.DocTitle_Col;
                int transid_InputRecordMaster_Col = InputRecordMasterExcelData.TRANSID_Col;

                int ver_InputRecordMaster_Col = InputRecordMasterExcelData_RLMU_Receive_1st.Ver_Col;
                int rev_InputRecordMaster_Col = InputRecordMasterExcelData_RLMU_Receive_1st.REV_Col;
                int purpose_InputRecordMaster_Col = InputRecordMasterExcelData_RLMU_Receive_1st.Purpose_Col;
                int status_InputRecordMaster_Col = InputRecordMasterExcelData_RLMU_Receive_1st.Status_Col;
                int date_InputRecordMaster_Col = InputRecordMasterExcelData_RLMU_Receive_1st.Date_Col;
                int timesCount_InputRecordMaster_Col = InputRecordMasterExcelData_RLMU_Receive_1st.TimesCount_Col;
                int package_InputRecordMaster_Col = InputRecordMasterExcelData_RLMU_Receive_1st.Package_Col;
                int discipline_InputRecordMaster_Col = InputRecordMasterExcelData_RLMU_Receive_1st.Discipline_Col;
                int redlineMarkup_InputRecordMaster_Col = InputRecordMasterExcelData_RLMU_Receive_1st.RedlineMarkup_Col;
                int packageStampStatus_InputRecordMaster_Col = InputRecordMasterExcelData_RLMU_Receive_1st.PackageStampStatus_Col;
                int date_Issue_InputRecordMaster_Col = -1;
                int modelName_InputRecordMaster_Col = -1;
                int nativeFileType_InputRecordMaster_Col = -1;
                int submittedFileType_InputRecordMaster_Col = -1;
                int atlasComment_InputRecordMaster_Col = InputRecordMasterExcelData_RLMU_Receive_1st.AtlasComment_Col;
                int updateStatus_InputRecordMaster_Col = InputRecordMasterExcelData_RLMU_Receive_1st.UpdateStatus_Col;
                int checkDocTitle_InputRecordMaster_Col = InputRecordMasterExcelData_RLMU_Receive_1st.CheckDocTitle_Col;

                // * Lấy về tổng số dòng của các file
                int totalRowsInput = wsInput.Dimension.End.Row;
                int totalRowsMaster = wsMaster.Dimension.End.Row;

                // *** Đưa data vào Master
                for (int row = InputExcelData.Start_Row; row <= totalRowsInput; row++)
                {
                    // Chỉ lấy dòng có STATUS = OK
                    if (wsInput.Cells[row, InputExcelData.UpdateStatus_Col].Text != "OK")
                    {
                        continue; // Bỏ qua dòng Duplicate
                    }

                    string docRef = wsInput.Cells[row, InputExcelData.DocRef_Col].Text;     // Lấy giá trị "Doc Ref" trong Input file
                    string docTitle = wsInput.Cells[row, InputExcelData.DocTitle_Col].Text;

                    string verValue = wsInput.Cells[row, InputExcelData.Ver_Col].Text;
                    string revValue = wsInput.Cells[row, InputExcelData.REV_Col].Text;
                    string purposeValue = wsInput.Cells[row, InputExcelData.Purpose_Col].Text;
                    string statusValue = wsInput.Cells[row, InputExcelData.Status_Col].Text;
                    string transIdValue = wsInput.Cells[row, InputExcelData.TRANSID_Col].Text;
                    string dateValue_str = wsInput.Cells[row, InputExcelData.Date_Col].Text;
                    string timesCountValue = wsInput.Cells[row, InputExcelData.TimesCount_Col].Text;
                    string packageValue = wsInput.Cells[row, InputExcelData.Package_Col].Text;
                    string disciplineValue = wsInput.Cells[row, InputExcelData.Discipline_Col].Text;
                    string redlineMarkupValue = wsInput.Cells[row, InputExcelData.RedlineMarkup_Col].Text;
                    string packageStampStatusValue = wsInput.Cells[row, InputExcelData.PackageStampStatus_Col].Text;
                    string nativeFileTypeValue = wsInput.Cells[row, InputExcelData.NativeFileType_Col].Text;
                    string submittedFileTypeValue = wsInput.Cells[row, InputExcelData.SubmittedFileType_Col].Text;
                    string modelNameValue = wsInput.Cells[row, InputExcelData.ModelName_Col].Text;
                    string atlasCommentValue = wsInput.Cells[row, InputExcelData.AtlasComment_Col].Text;
                    string updateStatusValue = wsInput.Cells[row, InputExcelData.UpdateStatus_Col].Text;
                    string checkDocTitleValue = wsInput.Cells[row, InputExcelData.CheckDocTitle_Col].Text;

                    if (string.IsNullOrEmpty(docRef)) { continue; }

                    for (int masterRow = MasterExcelData_Drawing.Start_Row; masterRow <= totalRowsMaster; masterRow++)
                    {
                        string allianceNo = wsMaster.Cells[masterRow, MasterExcelData_Drawing.Alliance_Col].Text;   // Giá trị "Alliance No." trong Master file
                        if (allianceNo == docRef)
                        {
                            // DOC TITLE
                            if (checkDocTitleValue.StartsWith("Need to update"))
                            {
                                wsMaster.Cells[masterRow, docTitle_Master_Col].Value = docTitle;
                            }
                            
                            // TIMES COUNT
                            wsMaster.Cells[masterRow, timesCount_Master_Col].Value = timesCountValue;
                            if (int.TryParse(timesCountValue, out int timesCountValue_int))
                            {
                                wsMaster.Cells[masterRow, timesCount_Master_Col].Value = timesCountValue_int;
                            }
                            // TRANS ID
                            wsMaster.Cells[masterRow, transid_Master_Col].Value = transIdValue;
                            // REV
                            wsMaster.Cells[masterRow, rev_Master_Col].Value = revValue;
                            if (int.TryParse(revValue, out int revValue_int))
                            {
                                wsMaster.Cells[masterRow, rev_Master_Col].Value = revValue_int;
                            }
                            // VER (ko có ở phần Submit)
                            wsMaster.Cells[masterRow, ver_Master_Col].Value = verValue;
                            if (int.TryParse(verValue, out int verValue_int))
                            {
                                wsMaster.Cells[masterRow, ver_Master_Col].Value = verValue_int;
                            }
                            // PURPOSE
                            wsMaster.Cells[masterRow, purpose_Master_Col].Value = purposeValue;
                            // DATE
                            DateTime dateValue_ddmmyyy;
                            if (DateTime.TryParseExact(dateValue_str, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out dateValue_ddmmyyy))
                            {
                                wsMaster.Cells[masterRow, date_Master_Col].Value = dateValue_ddmmyyy;
                                wsMaster.Cells[masterRow, date_Master_Col].Style.Numberformat.Format = "dd/MM/yyyy"; // Định dạng ô theo ngày tháng năm
                            }
                            else
                            {
                                wsMaster.Cells[masterRow, date_Master_Col].Value = dateValue_str; // Giữ nguyên nếu không parse được
                            }
                            // STATUS
                            wsMaster.Cells[masterRow, status_Master_Col].Value = statusValue;
                            // ATLAS COMMENT
                            wsMaster.Cells[masterRow, atlasComment_Master_Col].Value = atlasCommentValue;
                            // REDLINE MARKUP
                            wsMaster.Cells[masterRow, redlineMarkup_Master_Col].Value = redlineMarkupValue;
                            // PACKAGE STAMP STATUS
                            wsMaster.Cells[masterRow, packageStampStatus_Master_Col].Value = packageStampStatusValue;

                            break;
                        }
                    }
                }

                // *** Đưa data vào InputRecordMaster
                TransferToInputRecordMaster(wsInput, wsInputRecordMaster,
                    docRef_InputRecordMaster_Col,
                    docTitle_InputRecordMaster_Col,
                    ver_InputRecordMaster_Col,
                    rev_InputRecordMaster_Col,
                    purpose_InputRecordMaster_Col,
                    status_InputRecordMaster_Col,
                    transid_InputRecordMaster_Col,
                    date_InputRecordMaster_Col,
                    timesCount_InputRecordMaster_Col,
                    package_InputRecordMaster_Col,
                    discipline_InputRecordMaster_Col,
                    redlineMarkup_InputRecordMaster_Col,
                    packageStampStatus_InputRecordMaster_Col,
                    date_Issue_InputRecordMaster_Col,
                    modelName_InputRecordMaster_Col,
                    nativeFileType_InputRecordMaster_Col,
                    submittedFileType_InputRecordMaster_Col,
                    atlasComment_InputRecordMaster_Col,
                    updateStatus_InputRecordMaster_Col,
                    checkDocTitle_InputRecordMaster_Col,
                    dateTime_Checked,
                    out isInput_Modified);

                // *** Phần cuối: Lưu các file
                packageMaster.Save();
                packageInputRecordMaster.Save();
                if (isInput_Modified) { packageInput.Save(); }

                MessageBox.Show(activeForm, "Data transfer is done!", caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            #endregion
        }
        public static void Transfer_Drawing_Submit_1st_ToMaster()
        {
            #region
            string caption = "Transfer Drawing_Submit_1st to Master";
            string sheetName_InputRecorMaster = InputRecordMasterExcelData_Drawing_Submit_1st.sheetName;

            if (!IsValidExcelFiles_ForTransfer_MainTab(activeForm, caption, sheetName_InputRecorMaster)) { return; }

            //--------------------------------------------
            // *** Backup Master và InputRecordMaster trước khi transfer data
            // * Backup Master File
            BackupFileToBackupFolder(txtFilePath_Master, Path.Combine("_backup", "backup_Master"));

            // * Backup InputRecordMaster File
            BackupFileToBackupFolder(txtFilePath_InputRecordMaster, Path.Combine("_backup", "backup_" + sheetName_InputRecorMaster));

            //--------------------------------------------
            // *** Transfer data

            bool isInput_Modified = false;

            // * Lấy giá trị ngày giờ trong filename của InputData file
            string dateTime_Checked = "";
            string fileName_InputData = Path.GetFileName(txtFilePath_InputData);
            string[] parts = fileName_InputData.Split('_');

            if (parts.Length >= 4)
            {
                dateTime_Checked = parts[parts.Length - 2] + "_" + parts[parts.Length - 1].Replace(".xlsx", "");
            }

            using (var packageMaster = new ExcelPackage(new FileInfo(txtFilePath_Master)))
            using (var packageInput = new ExcelPackage(new FileInfo(txtFilePath_InputData)))
            using (var packageInputRecordMaster = new ExcelPackage(new FileInfo(txtFilePath_InputRecordMaster)))
            {
                // * Ktra có Sheet cần tìm trong Master file ko rồi mới bắt đầu
                var wsMaster = GetWorksheetByName(packageMaster, MasterExcelData_Drawing.SheetName);
                if (wsMaster == null)
                {
                    MessageBox.Show(activeForm, "Sheet " + MasterExcelData_Drawing.SheetName + " not found in the file " + Path.GetFileName(txtFilePath_Master), "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                var wsInput = packageInput.Workbook.Worksheets[InputData_WorksheetIndex_Main];
                var wsInputRecordMaster = GetWorksheetByName(packageInputRecordMaster, sheetName_InputRecorMaster);

                // * Thông tin các cột của Master file
                int docTitle_Master_Col = MasterExcelData_Drawing.DocTitle_Col;
                int timesCount_Master_Col = MasterExcelData_Drawing_FirstSubmission.TimesCount_Col;
                int transid_Master_Col = MasterExcelData_Drawing_FirstSubmission.TRANSID_Col;
                int rev_Master_Col = MasterExcelData_Drawing_FirstSubmission.REV_Col;
                int purpose_Master_Col = MasterExcelData_Drawing_FirstSubmission.Purpose_Col;
                int date_Master_Col = MasterExcelData_Drawing_FirstSubmission.Date_Col;
                int status_Master_Col = MasterExcelData_Drawing_FirstSubmission.Status_Col;
                int atlasComment_Master_Col = MasterExcelData_Drawing_FirstSubmission.AtlasComment_Col;

                // * Thông tin các cột của InputRecordMaster file
                int docRef_InputRecordMaster_Col = InputRecordMasterExcelData.DocRef_Col;
                int docTitle_InputRecordMaster_Col = InputRecordMasterExcelData.DocTitle_Col;
                int transid_InputRecordMaster_Col = InputRecordMasterExcelData.TRANSID_Col;

                int ver_InputRecordMaster_Col = InputRecordMasterExcelData_Drawing_Submit_1st.Ver_Col;
                int rev_InputRecordMaster_Col = InputRecordMasterExcelData_Drawing_Submit_1st.REV_Col;
                int purpose_InputRecordMaster_Col = InputRecordMasterExcelData_Drawing_Submit_1st.Purpose_Col;
                int status_InputRecordMaster_Col = InputRecordMasterExcelData_Drawing_Submit_1st.Status_Col;
                int date_InputRecordMaster_Col = InputRecordMasterExcelData_Drawing_Submit_1st.Date_Col;
                int timesCount_InputRecordMaster_Col = InputRecordMasterExcelData_Drawing_Submit_1st.TimesCount_Col;
                int package_InputRecordMaster_Col = InputRecordMasterExcelData_Drawing_Submit_1st.Package_Col;
                int discipline_InputRecordMaster_Col = InputRecordMasterExcelData_Drawing_Submit_1st.Discipline_Col;
                int date_Issue_InputRecordMaster_Col = -1;
                int redlineMarkup_InputRecordMaster_Col = -1;
                int packageStampStatus_InputRecordMaster_Col = -1;
                int nativeFileType_InputRecordMaster_Col = -1;
                int submittedFileType_InputRecordMaster_Col = -1;
                int modelName_InputRecordMaster_Col = InputRecordMasterExcelData_Drawing_Submit_1st.ModelName_Col;
                int atlasComment_InputRecordMaster_Col = InputRecordMasterExcelData_Drawing_Submit_1st.AtlasComment_Col;
                int updateStatus_InputRecordMaster_Col = InputRecordMasterExcelData_Drawing_Submit_1st.UpdateStatus_Col;
                int checkDocTitle_InputRecordMaster_Col = InputRecordMasterExcelData_Drawing_Submit_1st.CheckDocTitle_Col;

                // * Lấy về tổng số dòng của các file
                int totalRowsInput = wsInput.Dimension.End.Row;
                int totalRowsMaster = wsMaster.Dimension.End.Row;

                // *** Đưa data vào Master
                for (int row = InputExcelData.Start_Row; row <= totalRowsInput; row++)
                {
                    // Chỉ lấy dòng có STATUS = OK
                    if (wsInput.Cells[row, InputExcelData.UpdateStatus_Col].Text != "OK")
                    {
                        continue; // Bỏ qua dòng Duplicate
                    }

                    string docRef = wsInput.Cells[row, InputExcelData.DocRef_Col].Text;     // Lấy giá trị "Doc Ref" trong Input file
                    string docTitle = wsInput.Cells[row, InputExcelData.DocTitle_Col].Text;

                    string verValue = wsInput.Cells[row, InputExcelData.Ver_Col].Text;
                    string revValue = wsInput.Cells[row, InputExcelData.REV_Col].Text;
                    string purposeValue = wsInput.Cells[row, InputExcelData.Purpose_Col].Text;
                    string statusValue = wsInput.Cells[row, InputExcelData.Status_Col].Text;
                    string transIdValue = wsInput.Cells[row, InputExcelData.TRANSID_Col].Text;
                    string dateValue_str = wsInput.Cells[row, InputExcelData.Date_Col].Text;
                    string timesCountValue = wsInput.Cells[row, InputExcelData.TimesCount_Col].Text;
                    string packageValue = wsInput.Cells[row, InputExcelData.Package_Col].Text;
                    string disciplineValue = wsInput.Cells[row, InputExcelData.Discipline_Col].Text;
                    string nativeFileTypeValue = wsInput.Cells[row, InputExcelData.NativeFileType_Col].Text;
                    string submittedFileTypeValue = wsInput.Cells[row, InputExcelData.SubmittedFileType_Col].Text;
                    string modelNameValue = wsInput.Cells[row, InputExcelData.ModelName_Col].Text;
                    string atlasCommentValue = wsInput.Cells[row, InputExcelData.AtlasComment_Col].Text;
                    string updateStatusValue = wsInput.Cells[row, InputExcelData.UpdateStatus_Col].Text;
                    string checkDocTitleValue = wsInput.Cells[row, InputExcelData.CheckDocTitle_Col].Text;

                    if (string.IsNullOrEmpty(docRef)) { continue; }

                    for (int masterRow = MasterExcelData_Drawing.Start_Row; masterRow <= totalRowsMaster; masterRow++)
                    {
                        string allianceNo = wsMaster.Cells[masterRow, MasterExcelData_Drawing.Alliance_Col].Text;   // Giá trị "Alliance No." trong Master file
                        if (allianceNo == docRef)
                        {
                            // DOC TITLE
                            if (checkDocTitleValue.StartsWith("Need to update"))
                            {
                                wsMaster.Cells[masterRow, docTitle_Master_Col].Value = docTitle;
                            }

                            // TIMES COUNT
                            wsMaster.Cells[masterRow, timesCount_Master_Col].Value = timesCountValue;
                            if (int.TryParse(timesCountValue, out int timesCountValue_int))
                            {
                                wsMaster.Cells[masterRow, timesCount_Master_Col].Value = timesCountValue_int;
                            }
                            // TRANS ID
                            wsMaster.Cells[masterRow, transid_Master_Col].Value = transIdValue;
                            // REV
                            wsMaster.Cells[masterRow, rev_Master_Col].Value = revValue;
                            if (int.TryParse(revValue, out int revValue_int))
                            {
                                wsMaster.Cells[masterRow, rev_Master_Col].Value = revValue_int;
                            }
                            // VER (ko có ở phần Submit)
                            //wsMaster.Cells[masterRow, ver_Master_Col].Value = verValue;
                            // PURPOSE
                            wsMaster.Cells[masterRow, purpose_Master_Col].Value = purposeValue;
                            // DATE
                            DateTime dateValue_ddmmyyy;
                            if (DateTime.TryParseExact(dateValue_str, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out dateValue_ddmmyyy))
                            {
                                wsMaster.Cells[masterRow, date_Master_Col].Value = dateValue_ddmmyyy;
                                wsMaster.Cells[masterRow, date_Master_Col].Style.Numberformat.Format = "dd/MM/yyyy"; // Định dạng ô theo ngày tháng năm
                            }
                            else
                            {
                                wsMaster.Cells[masterRow, date_Master_Col].Value = dateValue_str; // Giữ nguyên nếu không parse được
                            }
                            // STATUS
                            wsMaster.Cells[masterRow, status_Master_Col].Value = statusValue;
                            // ATLAS COMMENT
                            wsMaster.Cells[masterRow, atlasComment_Master_Col].Value = atlasCommentValue;

                            break;
                        }
                    }
                }

                // *** Đưa data vào InputRecordMaster
                TransferToInputRecordMaster(wsInput, wsInputRecordMaster,
                    docRef_InputRecordMaster_Col,
                    docTitle_InputRecordMaster_Col,
                    ver_InputRecordMaster_Col,
                    rev_InputRecordMaster_Col,
                    purpose_InputRecordMaster_Col,
                    status_InputRecordMaster_Col,
                    transid_InputRecordMaster_Col,
                    date_InputRecordMaster_Col,
                    timesCount_InputRecordMaster_Col,
                    package_InputRecordMaster_Col,
                    discipline_InputRecordMaster_Col,
                    redlineMarkup_InputRecordMaster_Col,
                    packageStampStatus_InputRecordMaster_Col,
                    date_Issue_InputRecordMaster_Col,
                    modelName_InputRecordMaster_Col,
                    nativeFileType_InputRecordMaster_Col,
                    submittedFileType_InputRecordMaster_Col,
                    atlasComment_InputRecordMaster_Col,
                    updateStatus_InputRecordMaster_Col,
                    checkDocTitle_InputRecordMaster_Col,
                    dateTime_Checked,
                    out isInput_Modified);

                // *** Phần cuối: Lưu các file
                packageMaster.Save();
                packageInputRecordMaster.Save();
                if (isInput_Modified) { packageInput.Save(); }

                MessageBox.Show(activeForm, "Data transfer is done!", caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            #endregion
        }
        public static void Transfer_Drawing_RFI_1st_ToMaster()
        {
            #region
            string caption = "Transfer Drawing_RFI_1st to Master";
            string sheetName_InputRecorMaster = InputRecordMasterExcelData_Drawing_RFI_1st.sheetName;

            //--------------------------------------------
            // *** Backup Master và InputRecordMaster trước khi transfer data
            // * Backup Master File
            BackupFileToBackupFolder(txtFilePath_Master, Path.Combine("_backup", "backup_Master"));

            // * Backup InputRecordMaster File
            //BackupFileToBackupFolder(txtFilePath_InputRecordMaster, Path.Combine("_backup", "backup_" + sheetName_InputRecorMaster));

            //--------------------------------------------
            // *** Transfer data

            var processor = new RFIMasterProcessor();

            string rfiFilePath = txtFilePath_InputData;
            string masterFilePath = txtFilePath_Master;

            try
            {
                processor.ProcessRFIAndMasterFiles(rfiFilePath, masterFilePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ Lỗi: {ex.Message}");
            }
            
            MessageBox.Show(activeForm, "Data transfer is done!", caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
           

            #endregion
        }
        public static void Transfer_Drawing_RFI_1st_ToMaster_OLD()
        {
            #region
            /*
            string caption = "Transfer Drawing_RFI_1st to Master";
            string sheetName_InputRecorMaster = InputRecordMasterExcelData_Drawing_RFI_1st.sheetName;

            //if (!IsValidExcelFiles_ForTransfer_MainTab(activeForm, caption, sheetName_InputRecorMaster)) { return; }

            //--------------------------------------------
            // *** Backup Master và InputRecordMaster trước khi transfer data
            // * Backup Master File
            BackupFileToBackupFolder(txtFilePath_Master, Path.Combine("_backup", "backup_Master"));

            // * Backup InputRecordMaster File
            BackupFileToBackupFolder(txtFilePath_InputRecordMaster, Path.Combine("_backup", "backup_" + sheetName_InputRecorMaster));

            //--------------------------------------------
            // *** Transfer data

            bool isInput_Modified = false;

            // * Lấy giá trị ngày giờ trong filename của InputData file
            string dateTime_Checked = "";
            string fileName_InputData = Path.GetFileName(txtFilePath_InputData);
            string[] parts = fileName_InputData.Split('_');

            if (parts.Length >= 4)
            {
                dateTime_Checked = parts[parts.Length - 2] + "_" + parts[parts.Length - 1].Replace(".xlsx", "");
            }

            using (var packageMaster = new ExcelPackage(new FileInfo(txtFilePath_Master)))
            using (var packageInput = new ExcelPackage(new FileInfo(txtFilePath_InputData)))
            using (var packageInputRecordMaster = new ExcelPackage(new FileInfo(txtFilePath_InputRecordMaster)))
            {
                // * Ktra có Sheet cần tìm trong Master file ko rồi mới bắt đầu
                var wsMaster = GetWorksheetByName(packageMaster, MasterExcelData_Drawing.SheetName);
                if (wsMaster == null)
                {
                    MessageBox.Show(activeForm, "Sheet " + MasterExcelData_Drawing.SheetName + " not found in the file " + Path.GetFileName(txtFilePath_Master), "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                var wsInput = packageInput.Workbook.Worksheets[InputData_WorksheetIndex_Main];
                var wsInputRecordMaster = GetWorksheetByName(packageInputRecordMaster, sheetName_InputRecorMaster);

                // * Thông tin các cột của Master file
                int docTitle_Master_Col = MasterExcelData_Drawing.DocTitle_Col;
                int timesCount_Master_Col = MasterExcelData_Drawing_FirstRFI.TimesCount_Col;
                int number_Master_Col = MasterExcelData_Drawing_FirstRFI.No_Col;
                int dateRequest_Master_Col = MasterExcelData_Drawing_FirstRFI.Date_Col;
                int dateFeedback_Master_Col = MasterExcelData_Drawing_FirstRFI.DateFeedback_Col;
                int status_Master_Col = MasterExcelData_Drawing_FirstRFI.Status_Col;

                // * Thông tin các cột của InputRecordMaster file
                int number_InputRecordMaster_Col = InputRecordMasterExcelData_Drawing_RFI_1st.Number_Col;
                int docRef_InputRecordMaster_Col = InputRecordMasterExcelData_Drawing_RFI_1st.DocRef_Col;
                int docTitle_InputRecordMaster_Col = InputRecordMasterExcelData_Drawing_RFI_1st.DocTitle_Col;

                int drawingNo_InputRecordMaster_Col = InputRecordMasterExcelData_Drawing_RFI_1st.DrawingNo_Col;
                int dateRequest_InputRecordMaster_Col = InputRecordMasterExcelData_Drawing_RFI_1st.DateRequest_Col;
                int dateFeedback_InputRecordMaster_Col = InputRecordMasterExcelData_Drawing_RFI_1st.DateFeedback_Col;
                int statusRFI_InputRecordMaster_Col = InputRecordMasterExcelData_Drawing_RFI_1st.StatusRFI_Col;
                int timesCount_InputRecordMaster_Col = InputRecordMasterExcelData_Drawing_RFI_1st.TimesCount_Col;

                // * Lấy về tổng số dòng của các file
                int totalRowsInput = wsInput.Dimension.End.Row;
                int totalRowsMaster = wsMaster.Dimension.End.Row;

                // *** Đưa data vào Master
                for (int row = InputExcelData_RFI.Start_Row; row <= totalRowsInput; row++)
                {
                    string numberValue = wsInput.Cells[row, InputExcelData_RFI.No_Col].Text;
                    string dateRequestValue_str = wsInput.Cells[row, InputExcelData_RFI.DateRequest_Col].Text;
                    string drawingNoValue = wsInput.Cells[row, InputExcelData_RFI.DrawingNo_Col].Text;
                    string docRef = wsInput.Cells[row, InputExcelData_RFI.DocRef_Col].Text;     // Lấy giá trị "Doc Ref" trong Input file
                    string docTitle = wsInput.Cells[row, InputExcelData_RFI.DocTitle_Col].Text;
                    string dateFeedbackValue_str = wsInput.Cells[row, InputExcelData_RFI.DateFeedback_Col].Text;
                    string statusRFIValue = wsInput.Cells[row, InputExcelData_RFI.StatusRFI_Col].Text;
                    string timesCountValue = wsInput.Cells[row, InputExcelData_RFI.StatusRFI_Col].Text;

                    if (string.IsNullOrEmpty(docRef)) { continue; }

                    for (int masterRow = MasterExcelData_Drawing.Start_Row; masterRow <= totalRowsMaster; masterRow++)
                    {
                        string allianceNo = wsMaster.Cells[masterRow, MasterExcelData_Drawing.Alliance_Col].Text;   // Giá trị "Alliance No." trong Master file
                        if (allianceNo == docRef)
                        {
                            // TIMES COUNT
                            wsMaster.Cells[masterRow, timesCount_Master_Col].Value = timesCountValue;
                            if (int.TryParse(timesCountValue, out int timesCountValue_int))
                            {
                                wsMaster.Cells[masterRow, timesCount_Master_Col].Value = timesCountValue_int;
                            }

                            // NUMBER
                            wsMaster.Cells[masterRow, number_Master_Col].Value = numberValue;
                            
                            // DATE REQUEST
                            DateTime dateValue_ddmmyyy;
                            if (DateTime.TryParseExact(dateRequestValue_str, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out dateValue_ddmmyyy))
                            {
                                wsMaster.Cells[masterRow, dateRequest_Master_Col].Value = dateValue_ddmmyyy;
                                wsMaster.Cells[masterRow, dateRequest_Master_Col].Style.Numberformat.Format = "dd/MM/yyyy"; // Định dạng ô theo ngày tháng năm
                            }
                            else
                            {
                                wsMaster.Cells[masterRow, dateRequest_Master_Col].Value = dateRequestValue_str; // Giữ nguyên nếu không parse được
                            }
                            // DATE FEEDBACK
                            if (DateTime.TryParseExact(dateFeedbackValue_str, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out dateValue_ddmmyyy))
                            {
                                wsMaster.Cells[masterRow, dateFeedback_Master_Col].Value = dateValue_ddmmyyy;
                                wsMaster.Cells[masterRow, dateFeedback_Master_Col].Style.Numberformat.Format = "dd/MM/yyyy"; // Định dạng ô theo ngày tháng năm
                            }
                            else
                            {
                                wsMaster.Cells[masterRow, dateFeedback_Master_Col].Value = dateFeedbackValue_str; // Giữ nguyên nếu không parse được
                            }

                            // STATUS
                            wsMaster.Cells[masterRow, status_Master_Col].Value = statusRFIValue;

                            break;
                        }
                    }
                }

                // *** Đưa data vào InputRecordMaster
                TransferToInputRecordMaster_RFI(wsInput, wsInputRecordMaster,
                    number_InputRecordMaster_Col,
                    docRef_InputRecordMaster_Col,
                    docTitle_InputRecordMaster_Col,
                    drawingNo_InputRecordMaster_Col,
                    dateRequest_InputRecordMaster_Col,
                    dateFeedback_InputRecordMaster_Col,
                    statusRFI_InputRecordMaster_Col,
                    timesCount_InputRecordMaster_Col);

                // *** Phần cuối: Lưu các file
                packageMaster.Save();
                packageInputRecordMaster.Save();
                if (isInput_Modified) { packageInput.Save(); }

                MessageBox.Show(activeForm, "Data transfer is done!", caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            */
            #endregion
        }
        public static void Transfer_Drawing_Feedback_1st_ToMaster()
        {
            #region
            string caption = "Transfer Drawing_Feedback_1st to Master";
            string sheetName_InputRecorMaster = InputRecordMasterExcelData_Drawing_Feedback_1st.sheetName;

            if (!IsValidExcelFiles_ForTransfer_MainTab(activeForm, caption, sheetName_InputRecorMaster)) { return; }

            //--------------------------------------------
            // *** Backup Master và InputRecordMaster trước khi transfer data
            // * Backup Master File
            BackupFileToBackupFolder(txtFilePath_Master, Path.Combine("_backup", "backup_Master"));

            // * Backup InputRecordMaster File
            BackupFileToBackupFolder(txtFilePath_InputRecordMaster, Path.Combine("_backup", "backup_" + sheetName_InputRecorMaster));

            //--------------------------------------------
            // *** Transfer data

            bool isInput_Modified = false;

            // * Lấy giá trị ngày giờ trong filename của InputData file
            string dateTime_Checked = "";
            string fileName_InputData = Path.GetFileName(txtFilePath_InputData);
            string[] parts = fileName_InputData.Split('_');

            if (parts.Length >= 4)
            {
                dateTime_Checked = parts[parts.Length - 2] + "_" + parts[parts.Length - 1].Replace(".xlsx", "");
            }

            using (var packageMaster = new ExcelPackage(new FileInfo(txtFilePath_Master)))
            using (var packageInput = new ExcelPackage(new FileInfo(txtFilePath_InputData)))
            using (var packageInputRecordMaster = new ExcelPackage(new FileInfo(txtFilePath_InputRecordMaster)))
            {
                // * Ktra có Sheet cần tìm trong Master file ko rồi mới bắt đầu
                var wsMaster = GetWorksheetByName(packageMaster, MasterExcelData_Drawing.SheetName);
                if (wsMaster == null)
                {
                    MessageBox.Show(activeForm, "Sheet " + MasterExcelData_Drawing.SheetName + " not found in the file " + Path.GetFileName(txtFilePath_Master), "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                var wsInput = packageInput.Workbook.Worksheets[InputData_WorksheetIndex_Main];
                var wsInputRecordMaster = GetWorksheetByName(packageInputRecordMaster, sheetName_InputRecorMaster);

                // * Thông tin các cột của Master file
                int docTitle_Master_Col = MasterExcelData_Drawing.DocTitle_Col;
                int timesCount_Master_Col = MasterExcelData_Drawing_FirstFeedback.TimesCount_Col;
                int transid_Master_Col = MasterExcelData_Drawing_FirstFeedback.TRANSID_Col;
                int status_Master_Col = MasterExcelData_Drawing_FirstFeedback.Status_Col;
                int date_Master_Col = MasterExcelData_Drawing_FirstFeedback.Date_Col;
                int atlasComment_Master_Col = MasterExcelData_Drawing_FirstFeedback.AtlasComment_Col;

                // * Thông tin các cột của InputRecordMaster file
                int docRef_InputRecordMaster_Col = InputRecordMasterExcelData.DocRef_Col;
                int docTitle_InputRecordMaster_Col = InputRecordMasterExcelData.DocTitle_Col;
                int transid_InputRecordMaster_Col = InputRecordMasterExcelData.TRANSID_Col;

                int ver_InputRecordMaster_Col = InputRecordMasterExcelData_Drawing_Feedback_1st.Ver_Col;
                int rev_InputRecordMaster_Col = InputRecordMasterExcelData_Drawing_Feedback_1st.REV_Col;
                int purpose_InputRecordMaster_Col = InputRecordMasterExcelData_Drawing_Feedback_1st.Purpose_Col;
                int status_InputRecordMaster_Col = InputRecordMasterExcelData_Drawing_Feedback_1st.Status_Col;
                int date_InputRecordMaster_Col = InputRecordMasterExcelData_Drawing_Feedback_1st.Date_Col;
                int timesCount_InputRecordMaster_Col = InputRecordMasterExcelData_Drawing_Feedback_1st.TimesCount_Col;
                int package_InputRecordMaster_Col = InputRecordMasterExcelData_Drawing_Feedback_1st.Package_Col;
                int discipline_InputRecordMaster_Col = InputRecordMasterExcelData_Drawing_Feedback_1st.Discipline_Col;
                int date_Issue_InputRecordMaster_Col = -1;
                int redlineMarkup_InputRecordMaster_Col = -1;
                int packageStampStatus_InputRecordMaster_Col = -1;
                int nativeFileType_InputRecordMaster_Col = -1;
                int submittedFileType_InputRecordMaster_Col = -1;
                int modelName_InputRecordMaster_Col = InputRecordMasterExcelData_Drawing_Feedback_1st.ModelName_Col;
                int atlasComment_InputRecordMaster_Col = InputRecordMasterExcelData_Drawing_Feedback_1st.AtlasComment_Col;
                int updateStatus_InputRecordMaster_Col = InputRecordMasterExcelData_Drawing_Feedback_1st.UpdateStatus_Col;
                int checkDocTitle_InputRecordMaster_Col = InputRecordMasterExcelData_Drawing_Feedback_1st.CheckDocTitle_Col;

                // * Lấy về tổng số dòng của các file
                int totalRowsInput = wsInput.Dimension.End.Row;
                int totalRowsMaster = wsMaster.Dimension.End.Row;

                // *** Đưa data vào Master
                for (int row = InputExcelData.Start_Row; row <= totalRowsInput; row++)
                {
                    // Chỉ lấy dòng có STATUS = OK
                    if (wsInput.Cells[row, InputExcelData.UpdateStatus_Col].Text != "OK")
                    {
                        continue; // Bỏ qua dòng Duplicate
                    }

                    string docRef = wsInput.Cells[row, InputExcelData.DocRef_Col].Text;     // Lấy giá trị "Doc Ref" trong Input file
                    string docTitle = wsInput.Cells[row, InputExcelData.DocTitle_Col].Text;

                    string verValue = wsInput.Cells[row, InputExcelData.Ver_Col].Text;
                    string revValue = wsInput.Cells[row, InputExcelData.REV_Col].Text;
                    string purposeValue = wsInput.Cells[row, InputExcelData.Purpose_Col].Text;
                    string statusValue = wsInput.Cells[row, InputExcelData.Status_Col].Text;
                    string transIdValue = wsInput.Cells[row, InputExcelData.TRANSID_Col].Text;
                    string dateValue_str = wsInput.Cells[row, InputExcelData.Date_Col].Text;
                    string timesCountValue = wsInput.Cells[row, InputExcelData.TimesCount_Col].Text;
                    string packageValue = wsInput.Cells[row, InputExcelData.Package_Col].Text;
                    string disciplineValue = wsInput.Cells[row, InputExcelData.Discipline_Col].Text;
                    string nativeFileTypeValue = wsInput.Cells[row, InputExcelData.NativeFileType_Col].Text;
                    string submittedFileTypeValue = wsInput.Cells[row, InputExcelData.SubmittedFileType_Col].Text;
                    string modelNameValue = wsInput.Cells[row, InputExcelData.ModelName_Col].Text;
                    string atlasCommentValue = wsInput.Cells[row, InputExcelData.AtlasComment_Col].Text;
                    string updateStatusValue = wsInput.Cells[row, InputExcelData.UpdateStatus_Col].Text;
                    string checkDocTitleValue = wsInput.Cells[row, InputExcelData.CheckDocTitle_Col].Text;

                    if (string.IsNullOrEmpty(docRef)) { continue; }

                    for (int masterRow = MasterExcelData_Drawing.Start_Row; masterRow <= totalRowsMaster; masterRow++)
                    {
                        string allianceNo = wsMaster.Cells[masterRow, MasterExcelData_Drawing.Alliance_Col].Text;   // Giá trị "Alliance No." trong Master file
                        if (allianceNo == docRef)
                        {
                            // DOC TITLE
                            if (checkDocTitleValue.StartsWith("Need to update"))
                            {
                                wsMaster.Cells[masterRow, docTitle_Master_Col].Value = docTitle;
                            }

                            // TIMES COUNT
                            wsMaster.Cells[masterRow, timesCount_Master_Col].Value = timesCountValue;
                            if (int.TryParse(timesCountValue, out int timesCountValue_int))
                            {
                                wsMaster.Cells[masterRow, timesCount_Master_Col].Value = timesCountValue_int;
                            }
                            // TRANS ID
                            wsMaster.Cells[masterRow, transid_Master_Col].Value = transIdValue;
                            // STATUS
                            wsMaster.Cells[masterRow, status_Master_Col].Value = statusValue;
                            // DATE
                            DateTime dateValue_ddmmyyy;
                            if (DateTime.TryParseExact(dateValue_str, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out dateValue_ddmmyyy))
                            {
                                wsMaster.Cells[masterRow, date_Master_Col].Value = dateValue_ddmmyyy;
                                wsMaster.Cells[masterRow, date_Master_Col].Style.Numberformat.Format = "dd/MM/yyyy"; // Định dạng ô theo ngày tháng năm
                            }
                            else
                            {
                                wsMaster.Cells[masterRow, date_Master_Col].Value = dateValue_str; // Giữ nguyên nếu không parse được
                            }
                            // ATLAS COMMENT
                            wsMaster.Cells[masterRow, atlasComment_Master_Col].Value = atlasCommentValue;

                            break;
                        }
                    }
                }

                // *** Đưa data vào InputRecordMaster
                TransferToInputRecordMaster(wsInput, wsInputRecordMaster,
                    docRef_InputRecordMaster_Col,
                    docTitle_InputRecordMaster_Col,
                    ver_InputRecordMaster_Col,
                    rev_InputRecordMaster_Col,
                    purpose_InputRecordMaster_Col,
                    status_InputRecordMaster_Col,
                    transid_InputRecordMaster_Col,
                    date_InputRecordMaster_Col,
                    timesCount_InputRecordMaster_Col,
                    package_InputRecordMaster_Col,
                    discipline_InputRecordMaster_Col,
                    redlineMarkup_InputRecordMaster_Col,
                    packageStampStatus_InputRecordMaster_Col,
                    date_Issue_InputRecordMaster_Col,
                    modelName_InputRecordMaster_Col,
                    nativeFileType_InputRecordMaster_Col,
                    submittedFileType_InputRecordMaster_Col,
                    atlasComment_InputRecordMaster_Col,
                    updateStatus_InputRecordMaster_Col,
                    checkDocTitle_InputRecordMaster_Col,
                    dateTime_Checked,
                    out isInput_Modified);

                // *** Phần cuối: Lưu các file
                packageMaster.Save();
                packageInputRecordMaster.Save();
                if (isInput_Modified) { packageInput.Save(); }

                MessageBox.Show(activeForm, "Data transfer is done!", caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            #endregion
        }

        //-------------------------------------------------------
        // 02. DRAWING - FOR FINAL
        public static void Transfer_RLMU_Receive_Final_ToMaster()
        {
            #region
            string caption = "Transfer RLMU_Receive_Final to Master";
            string sheetName_InputRecorMaster = InputRecordMasterExcelData_RLMU_Receive_Final.sheetName;

            if (!IsValidExcelFiles_ForTransfer_MainTab(activeForm, caption, sheetName_InputRecorMaster)) { return; }

            //--------------------------------------------
            // *** Backup Master và InputRecordMaster trước khi transfer data
            // * Backup Master File
            BackupFileToBackupFolder(txtFilePath_Master, Path.Combine("_backup", "backup_Master"));

            // * Backup InputRecordMaster File
            BackupFileToBackupFolder(txtFilePath_InputRecordMaster, Path.Combine("_backup", "backup_" + sheetName_InputRecorMaster));

            //--------------------------------------------
            // *** Transfer data
            bool isInput_Modified = false;

            // * Lấy giá trị ngày giờ trong filename của InputData file
            string dateTime_Checked = "";
            string fileName_InputData = Path.GetFileName(txtFilePath_InputData);
            string[] parts = fileName_InputData.Split('_');

            if (parts.Length >= 4)
            {
                dateTime_Checked = parts[parts.Length - 2] + "_" + parts[parts.Length - 1].Replace(".xlsx", "");
            }

            using (var packageMaster = new ExcelPackage(new FileInfo(txtFilePath_Master)))
            using (var packageInput = new ExcelPackage(new FileInfo(txtFilePath_InputData)))
            using (var packageInputRecordMaster = new ExcelPackage(new FileInfo(txtFilePath_InputRecordMaster)))
            {
                // * Ktra có Sheet cần tìm trong Master file ko rồi mới bắt đầu
                var wsMaster = GetWorksheetByName(packageMaster, MasterExcelData_Drawing.SheetName);
                if (wsMaster == null)
                {
                    MessageBox.Show(activeForm, "Sheet " + MasterExcelData_Drawing.SheetName + " not found in the file " + Path.GetFileName(txtFilePath_Master), "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                var wsInput = packageInput.Workbook.Worksheets[InputData_WorksheetIndex_Main];
                var wsInputRecordMaster = GetWorksheetByName(packageInputRecordMaster, sheetName_InputRecorMaster);

                // * Thông tin các cột của Master file
                int docTitle_Master_Col = MasterExcelData_Drawing.DocTitle_Col;
                int timesCount_Master_Col = MasterExcelData_Drawing_FinalReceive.TimesCount_Col;
                int transid_Master_Col = MasterExcelData_Drawing_FinalReceive.TRANSID_Col;
                int rev_Master_Col = MasterExcelData_Drawing_FinalReceive.REV_Col;
                int ver_Master_Col = MasterExcelData_Drawing_FinalReceive.Ver_Col;
                int purpose_Master_Col = MasterExcelData_Drawing_FinalReceive.Purpose_Col;
                int date_Master_Col = MasterExcelData_Drawing_FinalReceive.Date_Col;
                int status_Master_Col = MasterExcelData_Drawing_FinalReceive.Status_Col;
                int atlasComment_Master_Col = MasterExcelData_Drawing_FinalReceive.AtlasComment_Col;
                int redlineMarkup_Master_Col = MasterExcelData_Drawing_FinalReceive.RedlineMarkup_Col;
                int packageStampStatus_Master_Col = MasterExcelData_Drawing_FinalReceive.PackageStampStatus_Col;

                // * Thông tin các cột của InputRecordMaster file
                int docRef_InputRecordMaster_Col = InputRecordMasterExcelData.DocRef_Col;
                int docTitle_InputRecordMaster_Col = InputRecordMasterExcelData.DocTitle_Col;
                int transid_InputRecordMaster_Col = InputRecordMasterExcelData.TRANSID_Col;

                int ver_InputRecordMaster_Col = InputRecordMasterExcelData_RLMU_Receive_Final.Ver_Col;
                int rev_InputRecordMaster_Col = InputRecordMasterExcelData_RLMU_Receive_Final.REV_Col;
                int purpose_InputRecordMaster_Col = InputRecordMasterExcelData_RLMU_Receive_Final.Purpose_Col;
                int status_InputRecordMaster_Col = InputRecordMasterExcelData_RLMU_Receive_Final.Status_Col;
                int date_InputRecordMaster_Col = InputRecordMasterExcelData_RLMU_Receive_Final.Date_Col;
                int timesCount_InputRecordMaster_Col = InputRecordMasterExcelData_RLMU_Receive_Final.TimesCount_Col;
                int package_InputRecordMaster_Col = InputRecordMasterExcelData_RLMU_Receive_Final.Package_Col;
                int discipline_InputRecordMaster_Col = InputRecordMasterExcelData_RLMU_Receive_Final.Discipline_Col;
                int redlineMarkup_InputRecordMaster_Col = InputRecordMasterExcelData_RLMU_Receive_Final.RedlineMarkup_Col;
                int packageStampStatus_InputRecordMaster_Col = InputRecordMasterExcelData_RLMU_Receive_Final.PackageStampStatus_Col;
                int date_Issue_InputRecordMaster_Col = -1;
                int modelName_InputRecordMaster_Col = -1;
                int nativeFileType_InputRecordMaster_Col = -1;
                int submittedFileType_InputRecordMaster_Col = -1;
                int atlasComment_InputRecordMaster_Col = InputRecordMasterExcelData_RLMU_Receive_Final.AtlasComment_Col;
                int updateStatus_InputRecordMaster_Col = InputRecordMasterExcelData_RLMU_Receive_Final.UpdateStatus_Col;
                int checkDocTitle_InputRecordMaster_Col = InputRecordMasterExcelData_RLMU_Receive_Final.CheckDocTitle_Col;

                // * Lấy về tổng số dòng của các file
                int totalRowsInput = wsInput.Dimension.End.Row;
                int totalRowsMaster = wsMaster.Dimension.End.Row;

                // *** Đưa data vào Master
                for (int row = InputExcelData.Start_Row; row <= totalRowsInput; row++)
                {
                    // Chỉ lấy dòng có STATUS = OK
                    if (wsInput.Cells[row, InputExcelData.UpdateStatus_Col].Text != "OK")
                    {
                        continue; // Bỏ qua dòng Duplicate
                    }

                    string docRef = wsInput.Cells[row, InputExcelData.DocRef_Col].Text;     // Lấy giá trị "Doc Ref" trong Input file
                    string docTitle = wsInput.Cells[row, InputExcelData.DocTitle_Col].Text;

                    string verValue = wsInput.Cells[row, InputExcelData.Ver_Col].Text;
                    string revValue = wsInput.Cells[row, InputExcelData.REV_Col].Text;
                    string purposeValue = wsInput.Cells[row, InputExcelData.Purpose_Col].Text;
                    string statusValue = wsInput.Cells[row, InputExcelData.Status_Col].Text;
                    string transIdValue = wsInput.Cells[row, InputExcelData.TRANSID_Col].Text;
                    string dateValue_str = wsInput.Cells[row, InputExcelData.Date_Col].Text;
                    string timesCountValue = wsInput.Cells[row, InputExcelData.TimesCount_Col].Text;
                    string packageValue = wsInput.Cells[row, InputExcelData.Package_Col].Text;
                    string disciplineValue = wsInput.Cells[row, InputExcelData.Discipline_Col].Text;
                    string redlineMarkupValue = wsInput.Cells[row, InputExcelData.RedlineMarkup_Col].Text;
                    string packageStampStatusValue = wsInput.Cells[row, InputExcelData.PackageStampStatus_Col].Text;
                    string nativeFileTypeValue = wsInput.Cells[row, InputExcelData.NativeFileType_Col].Text;
                    string submittedFileTypeValue = wsInput.Cells[row, InputExcelData.SubmittedFileType_Col].Text;
                    string modelNameValue = wsInput.Cells[row, InputExcelData.ModelName_Col].Text;
                    string atlasCommentValue = wsInput.Cells[row, InputExcelData.AtlasComment_Col].Text;
                    string updateStatusValue = wsInput.Cells[row, InputExcelData.UpdateStatus_Col].Text;
                    string checkDocTitleValue = wsInput.Cells[row, InputExcelData.CheckDocTitle_Col].Text;

                    if (string.IsNullOrEmpty(docRef)) { continue; }

                    for (int masterRow = MasterExcelData_Drawing.Start_Row; masterRow <= totalRowsMaster; masterRow++)
                    {
                        string allianceNo = wsMaster.Cells[masterRow, MasterExcelData_Drawing.Alliance_Col].Text;   // Giá trị "Alliance No." trong Master file
                        if (allianceNo == docRef)
                        {
                            // DOC TITLE
                            if (checkDocTitleValue.StartsWith("Need to update"))
                            {
                                wsMaster.Cells[masterRow, docTitle_Master_Col].Value = docTitle;
                            }

                            // TIMES COUNT
                            wsMaster.Cells[masterRow, timesCount_Master_Col].Value = timesCountValue;
                            if (int.TryParse(timesCountValue, out int timesCountValue_int))
                            {
                                wsMaster.Cells[masterRow, timesCount_Master_Col].Value = timesCountValue_int;
                            }
                            // TRANS ID
                            wsMaster.Cells[masterRow, transid_Master_Col].Value = transIdValue;
                            // REV
                            wsMaster.Cells[masterRow, rev_Master_Col].Value = revValue;
                            if (int.TryParse(revValue, out int revValue_int))
                            {
                                wsMaster.Cells[masterRow, rev_Master_Col].Value = revValue_int;
                            }
                            // VER (ko có ở phần Submit)
                            wsMaster.Cells[masterRow, ver_Master_Col].Value = verValue;
                            if (int.TryParse(verValue, out int verValue_int))
                            {
                                wsMaster.Cells[masterRow, ver_Master_Col].Value = verValue_int;
                            }
                            // PURPOSE
                            wsMaster.Cells[masterRow, purpose_Master_Col].Value = purposeValue;
                            // DATE
                            DateTime dateValue_ddmmyyy;
                            if (DateTime.TryParseExact(dateValue_str, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out dateValue_ddmmyyy))
                            {
                                wsMaster.Cells[masterRow, date_Master_Col].Value = dateValue_ddmmyyy;
                                wsMaster.Cells[masterRow, date_Master_Col].Style.Numberformat.Format = "dd/MM/yyyy"; // Định dạng ô theo ngày tháng năm
                            }
                            else
                            {
                                wsMaster.Cells[masterRow, date_Master_Col].Value = dateValue_str; // Giữ nguyên nếu không parse được
                            }
                            // STATUS
                            wsMaster.Cells[masterRow, status_Master_Col].Value = statusValue;
                            // ATLAS COMMENT
                            wsMaster.Cells[masterRow, atlasComment_Master_Col].Value = atlasCommentValue;
                            // REDLINE MARKUP
                            wsMaster.Cells[masterRow, redlineMarkup_Master_Col].Value = redlineMarkupValue;
                            // PACKAGE STAMP STATUS
                            wsMaster.Cells[masterRow, packageStampStatus_Master_Col].Value = packageStampStatusValue;

                            break;
                        }
                    }
                }

                // *** Đưa data vào InputRecordMaster
                TransferToInputRecordMaster(wsInput, wsInputRecordMaster,
                    docRef_InputRecordMaster_Col,
                    docTitle_InputRecordMaster_Col,
                    ver_InputRecordMaster_Col,
                    rev_InputRecordMaster_Col,
                    purpose_InputRecordMaster_Col,
                    status_InputRecordMaster_Col,
                    transid_InputRecordMaster_Col,
                    date_InputRecordMaster_Col,
                    timesCount_InputRecordMaster_Col,
                    package_InputRecordMaster_Col,
                    discipline_InputRecordMaster_Col,
                    redlineMarkup_InputRecordMaster_Col,
                    packageStampStatus_InputRecordMaster_Col,
                    date_Issue_InputRecordMaster_Col,
                    modelName_InputRecordMaster_Col,
                    nativeFileType_InputRecordMaster_Col,
                    submittedFileType_InputRecordMaster_Col,
                    atlasComment_InputRecordMaster_Col,
                    updateStatus_InputRecordMaster_Col,
                    checkDocTitle_InputRecordMaster_Col,
                    dateTime_Checked,
                    out isInput_Modified);

                // *** Phần cuối: Lưu các file
                packageMaster.Save();
                packageInputRecordMaster.Save();
                if (isInput_Modified) { packageInput.Save(); }

                MessageBox.Show(activeForm, "Data transfer is done!", caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            #endregion
        }
        public static void Transfer_Drawing_Submit_Final_ToMaster()
        {
            #region
            string caption = "Transfer Drawing_Submit_Final to Master";
            string sheetName_InputRecorMaster = InputRecordMasterExcelData_Drawing_Submit_Final.sheetName;

            if (!IsValidExcelFiles_ForTransfer_MainTab(activeForm, caption, sheetName_InputRecorMaster)) { return; }

            //--------------------------------------------
            // *** Backup Master và InputRecordMaster trước khi transfer data
            // * Backup Master File
            BackupFileToBackupFolder(txtFilePath_Master, Path.Combine("_backup", "backup_Master"));

            // * Backup InputRecordMaster File
            BackupFileToBackupFolder(txtFilePath_InputRecordMaster, Path.Combine("_backup", "backup_" + sheetName_InputRecorMaster));

            //--------------------------------------------
            // *** Transfer data

            bool isInput_Modified = false;

            // * Lấy giá trị ngày giờ trong filename của InputData file
            string dateTime_Checked = "";
            string fileName_InputData = Path.GetFileName(txtFilePath_InputData);
            string[] parts = fileName_InputData.Split('_');

            if (parts.Length >= 4)
            {
                dateTime_Checked = parts[parts.Length - 2] + "_" + parts[parts.Length - 1].Replace(".xlsx", "");
            }

            using (var packageMaster = new ExcelPackage(new FileInfo(txtFilePath_Master)))
            using (var packageInput = new ExcelPackage(new FileInfo(txtFilePath_InputData)))
            using (var packageInputRecordMaster = new ExcelPackage(new FileInfo(txtFilePath_InputRecordMaster)))
            {
                // * Ktra có Sheet cần tìm trong Master file ko rồi mới bắt đầu
                var wsMaster = GetWorksheetByName(packageMaster, MasterExcelData_Drawing.SheetName);
                if (wsMaster == null)
                {
                    MessageBox.Show(activeForm, "Sheet " + MasterExcelData_Drawing.SheetName + " not found in the file " + Path.GetFileName(txtFilePath_Master), "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                var wsInput = packageInput.Workbook.Worksheets[InputData_WorksheetIndex_Main];
                var wsInputRecordMaster = GetWorksheetByName(packageInputRecordMaster, sheetName_InputRecorMaster);

                // * Thông tin các cột của Master file
                int docTitle_Master_Col = MasterExcelData_Drawing.DocTitle_Col;
                int timesCount_Master_Col = MasterExcelData_Drawing_FinalSubmission.TimesCount_Col;
                int transid_Master_Col = MasterExcelData_Drawing_FinalSubmission.TRANSID_Col;
                int rev_Master_Col = MasterExcelData_Drawing_FinalSubmission.REV_Col;
                int purpose_Master_Col = MasterExcelData_Drawing_FinalSubmission.Purpose_Col;
                int date_Master_Col = MasterExcelData_Drawing_FinalSubmission.Date_Col;
                int status_Master_Col = MasterExcelData_Drawing_FinalSubmission.Status_Col;
                int atlasComment_Master_Col = MasterExcelData_Drawing_FinalSubmission.AtlasComment_Col;

                // * Thông tin các cột của InputRecordMaster file
                int docRef_InputRecordMaster_Col = InputRecordMasterExcelData.DocRef_Col;
                int docTitle_InputRecordMaster_Col = InputRecordMasterExcelData.DocTitle_Col;
                int transid_InputRecordMaster_Col = InputRecordMasterExcelData.TRANSID_Col;

                int ver_InputRecordMaster_Col = InputRecordMasterExcelData_Drawing_Submit_Final.Ver_Col;
                int rev_InputRecordMaster_Col = InputRecordMasterExcelData_Drawing_Submit_Final.REV_Col;
                int purpose_InputRecordMaster_Col = InputRecordMasterExcelData_Drawing_Submit_Final.Purpose_Col;
                int status_InputRecordMaster_Col = InputRecordMasterExcelData_Drawing_Submit_Final.Status_Col;
                int date_InputRecordMaster_Col = InputRecordMasterExcelData_Drawing_Submit_Final.Date_Col;
                int timesCount_InputRecordMaster_Col = InputRecordMasterExcelData_Drawing_Submit_Final.TimesCount_Col;
                int package_InputRecordMaster_Col = InputRecordMasterExcelData_Drawing_Submit_Final.Package_Col;
                int discipline_InputRecordMaster_Col = InputRecordMasterExcelData_Drawing_Submit_Final.Discipline_Col;
                int redlineMarkup_InputRecordMaster_Col = -1;
                int packageStampStatus_InputRecordMaster_Col = -1;
                int date_Issue_InputRecordMaster_Col = -1;
                int modelName_InputRecordMaster_Col = InputRecordMasterExcelData_Drawing_Submit_Final.ModelName_Col;
                int nativeFileType_InputRecordMaster_Col = -1;
                int submittedFileType_InputRecordMaster_Col = -1;
                int atlasComment_InputRecordMaster_Col = InputRecordMasterExcelData_Drawing_Submit_Final.AtlasComment_Col;
                int updateStatus_InputRecordMaster_Col = InputRecordMasterExcelData_Drawing_Submit_Final.UpdateStatus_Col;
                int checkDocTitle_InputRecordMaster_Col = InputRecordMasterExcelData_Drawing_Submit_Final.CheckDocTitle_Col;

                // * Lấy về tổng số dòng của các file
                int totalRowsInput = wsInput.Dimension.End.Row;
                int totalRowsMaster = wsMaster.Dimension.End.Row;

                // *** Đưa data vào Master
                for (int row = InputExcelData.Start_Row; row <= totalRowsInput; row++)
                {
                    // Chỉ lấy dòng có STATUS = OK
                    if (wsInput.Cells[row, InputExcelData.UpdateStatus_Col].Text != "OK")
                    {
                        continue; // Bỏ qua dòng Duplicate
                    }

                    string docRef = wsInput.Cells[row, InputExcelData.DocRef_Col].Text;     // Lấy giá trị "Doc Ref" trong Input file
                    string docTitle = wsInput.Cells[row, InputExcelData.DocTitle_Col].Text;

                    string verValue = wsInput.Cells[row, InputExcelData.Ver_Col].Text;
                    string revValue = wsInput.Cells[row, InputExcelData.REV_Col].Text;
                    string purposeValue = wsInput.Cells[row, InputExcelData.Purpose_Col].Text;
                    string statusValue = wsInput.Cells[row, InputExcelData.Status_Col].Text;
                    string transIdValue = wsInput.Cells[row, InputExcelData.TRANSID_Col].Text;
                    string dateValue_str = wsInput.Cells[row, InputExcelData.Date_Col].Text;
                    string timesCountValue = wsInput.Cells[row, InputExcelData.TimesCount_Col].Text;
                    string packageValue = wsInput.Cells[row, InputExcelData.Package_Col].Text;
                    string disciplineValue = wsInput.Cells[row, InputExcelData.Discipline_Col].Text;
                    string nativeFileTypeValue = wsInput.Cells[row, InputExcelData.NativeFileType_Col].Text;
                    string submittedFileTypeValue = wsInput.Cells[row, InputExcelData.SubmittedFileType_Col].Text;
                    string modelNameValue = wsInput.Cells[row, InputExcelData.ModelName_Col].Text;
                    string atlasCommentValue = wsInput.Cells[row, InputExcelData.AtlasComment_Col].Text;
                    string updateStatusValue = wsInput.Cells[row, InputExcelData.UpdateStatus_Col].Text;
                    string checkDocTitleValue = wsInput.Cells[row, InputExcelData.CheckDocTitle_Col].Text;

                    if (string.IsNullOrEmpty(docRef)) { continue; }

                    for (int masterRow = MasterExcelData_Drawing.Start_Row; masterRow <= totalRowsMaster; masterRow++)
                    {
                        string allianceNo = wsMaster.Cells[masterRow, MasterExcelData_Drawing.Alliance_Col].Text;   // Giá trị "Alliance No." trong Master file
                        if (allianceNo == docRef)
                        {
                            // DOC TITLE
                            if (checkDocTitleValue.StartsWith("Need to update"))
                            {
                                wsMaster.Cells[masterRow, docTitle_Master_Col].Value = docTitle;
                            }

                            // TIMES COUNT
                            wsMaster.Cells[masterRow, timesCount_Master_Col].Value = timesCountValue;
                            if (int.TryParse(timesCountValue, out int timesCountValue_int))
                            {
                                wsMaster.Cells[masterRow, timesCount_Master_Col].Value = timesCountValue_int;
                            }
                            // TRANS ID
                            wsMaster.Cells[masterRow, transid_Master_Col].Value = transIdValue;
                            // REV
                            wsMaster.Cells[masterRow, rev_Master_Col].Value = revValue;
                            if (int.TryParse(revValue, out int revValue_int))
                            {
                                wsMaster.Cells[masterRow, rev_Master_Col].Value = revValue_int;
                            }
                            // VER (ko có ở phần Submit)
                            //wsMaster.Cells[masterRow, ver_Master_Col].Value = verValue;
                            // PURPOSE
                            wsMaster.Cells[masterRow, purpose_Master_Col].Value = purposeValue;
                            // DATE
                            DateTime dateValue_ddmmyyy;
                            if (DateTime.TryParseExact(dateValue_str, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out dateValue_ddmmyyy))
                            {
                                wsMaster.Cells[masterRow, date_Master_Col].Value = dateValue_ddmmyyy;
                                wsMaster.Cells[masterRow, date_Master_Col].Style.Numberformat.Format = "dd/MM/yyyy"; // Định dạng ô theo ngày tháng năm
                            }
                            else
                            {
                                wsMaster.Cells[masterRow, date_Master_Col].Value = dateValue_str; // Giữ nguyên nếu không parse được
                            }
                            // STATUS
                            wsMaster.Cells[masterRow, status_Master_Col].Value = statusValue;
                            // ATLAS COMMENT
                            wsMaster.Cells[masterRow, atlasComment_Master_Col].Value = atlasCommentValue;

                            break;
                        }
                    }
                }

                // *** Đưa data vào InputRecordMaster
                TransferToInputRecordMaster(wsInput, wsInputRecordMaster,
                    docRef_InputRecordMaster_Col,
                    docTitle_InputRecordMaster_Col,
                    ver_InputRecordMaster_Col,
                    rev_InputRecordMaster_Col,
                    purpose_InputRecordMaster_Col,
                    status_InputRecordMaster_Col,
                    transid_InputRecordMaster_Col,
                    date_InputRecordMaster_Col,
                    timesCount_InputRecordMaster_Col,
                    package_InputRecordMaster_Col,
                    discipline_InputRecordMaster_Col,
                    redlineMarkup_InputRecordMaster_Col,
                    packageStampStatus_InputRecordMaster_Col,
                    date_Issue_InputRecordMaster_Col,
                    modelName_InputRecordMaster_Col,
                    nativeFileType_InputRecordMaster_Col,
                    submittedFileType_InputRecordMaster_Col,
                    atlasComment_InputRecordMaster_Col,
                    updateStatus_InputRecordMaster_Col,
                    checkDocTitle_InputRecordMaster_Col,
                    dateTime_Checked,
                    out isInput_Modified);

                // *** Phần cuối: Lưu các file
                packageMaster.Save();
                packageInputRecordMaster.Save();
                if (isInput_Modified) { packageInput.Save(); }

                MessageBox.Show(activeForm, "Data transfer is done!", caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            #endregion
        }
        public static void Transfer_Drawing_RFI_Final_ToMaster()
        {

        }
        public static void Transfer_Drawing_Feedback_Final_ToMaster()
        {
            #region
            string caption = "Transfer Drawing_Feedback_Final to Master";
            string sheetName_InputRecorMaster = InputRecordMasterExcelData_Drawing_Feedback_Final.sheetName;

            if (!IsValidExcelFiles_ForTransfer_MainTab(activeForm, caption, sheetName_InputRecorMaster)) { return; }

            //--------------------------------------------
            // *** Backup Master và InputRecordMaster trước khi transfer data
            // * Backup Master File
            BackupFileToBackupFolder(txtFilePath_Master, Path.Combine("_backup", "backup_Master"));

            // * Backup InputRecordMaster File
            BackupFileToBackupFolder(txtFilePath_InputRecordMaster, Path.Combine("_backup", "backup_" + sheetName_InputRecorMaster));

            //--------------------------------------------
            // *** Transfer data

            bool isInput_Modified = false;

            // * Lấy giá trị ngày giờ trong filename của InputData file
            string dateTime_Checked = "";
            string fileName_InputData = Path.GetFileName(txtFilePath_InputData);
            string[] parts = fileName_InputData.Split('_');

            if (parts.Length >= 4)
            {
                dateTime_Checked = parts[parts.Length - 2] + "_" + parts[parts.Length - 1].Replace(".xlsx", "");
            }

            using (var packageMaster = new ExcelPackage(new FileInfo(txtFilePath_Master)))
            using (var packageInput = new ExcelPackage(new FileInfo(txtFilePath_InputData)))
            using (var packageInputRecordMaster = new ExcelPackage(new FileInfo(txtFilePath_InputRecordMaster)))
            {
                // * Ktra có Sheet cần tìm trong Master file ko rồi mới bắt đầu
                var wsMaster = GetWorksheetByName(packageMaster, MasterExcelData_Drawing.SheetName);
                if (wsMaster == null)
                {
                    MessageBox.Show(activeForm, "Sheet " + MasterExcelData_Drawing.SheetName + " not found in the file " + Path.GetFileName(txtFilePath_Master), "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                var wsInput = packageInput.Workbook.Worksheets[InputData_WorksheetIndex_Main];
                var wsInputRecordMaster = GetWorksheetByName(packageInputRecordMaster, sheetName_InputRecorMaster);

                // * Thông tin các cột của Master file
                int docTitle_Master_Col = MasterExcelData_Drawing.DocTitle_Col;
                int timesCount_Master_Col = MasterExcelData_Drawing_FinalFeedback.TimesCount_Col;
                int transid_Master_Col = MasterExcelData_Drawing_FinalFeedback.TRANSID_Col;
                int status_Master_Col = MasterExcelData_Drawing_FinalFeedback.Status_Col;
                int date_Master_Col = MasterExcelData_Drawing_FinalFeedback.Date_Col;
                int atlasComment_Master_Col = MasterExcelData_Drawing_FinalFeedback.AtlasComment_Col;

                // * Thông tin các cột của InputRecordMaster file
                int docRef_InputRecordMaster_Col = InputRecordMasterExcelData.DocRef_Col;
                int docTitle_InputRecordMaster_Col = InputRecordMasterExcelData.DocTitle_Col;
                int transid_InputRecordMaster_Col = InputRecordMasterExcelData.TRANSID_Col;

                int ver_InputRecordMaster_Col = InputRecordMasterExcelData_Drawing_Feedback_Final.Ver_Col;
                int rev_InputRecordMaster_Col = InputRecordMasterExcelData_Drawing_Feedback_Final.REV_Col;
                int purpose_InputRecordMaster_Col = InputRecordMasterExcelData_Drawing_Feedback_Final.Purpose_Col;
                int status_InputRecordMaster_Col = InputRecordMasterExcelData_Drawing_Feedback_Final.Status_Col;
                int date_InputRecordMaster_Col = InputRecordMasterExcelData_Drawing_Feedback_Final.Date_Col;
                int timesCount_InputRecordMaster_Col = InputRecordMasterExcelData_Drawing_Feedback_Final.TimesCount_Col;
                int package_InputRecordMaster_Col = InputRecordMasterExcelData_Drawing_Feedback_Final.Package_Col;
                int discipline_InputRecordMaster_Col = InputRecordMasterExcelData_Drawing_Feedback_Final.Discipline_Col;
                int redlineMarkup_InputRecordMaster_Col = -1;
                int packageStampStatus_InputRecordMaster_Col = -1;
                int date_Issue_InputRecordMaster_Col = -1;
                int modelName_InputRecordMaster_Col = InputRecordMasterExcelData_Drawing_Feedback_Final.ModelName_Col;
                int nativeFileType_InputRecordMaster_Col = -1;
                int submittedFileType_InputRecordMaster_Col = -1;
                int atlasComment_InputRecordMaster_Col = InputRecordMasterExcelData_Drawing_Feedback_Final.AtlasComment_Col;
                int updateStatus_InputRecordMaster_Col = InputRecordMasterExcelData_Drawing_Feedback_Final.UpdateStatus_Col;
                int checkDocTitle_InputRecordMaster_Col = InputRecordMasterExcelData_Drawing_Feedback_Final.CheckDocTitle_Col;

                // * Lấy về tổng số dòng của các file
                int totalRowsInput = wsInput.Dimension.End.Row;
                int totalRowsMaster = wsMaster.Dimension.End.Row;

                // *** Đưa data vào Master
                for (int row = InputExcelData.Start_Row; row <= totalRowsInput; row++)
                {
                    // Chỉ lấy dòng có STATUS = OK
                    if (wsInput.Cells[row, InputExcelData.UpdateStatus_Col].Text != "OK")
                    {
                        continue; // Bỏ qua dòng Duplicate
                    }

                    string docRef = wsInput.Cells[row, InputExcelData.DocRef_Col].Text;     // Lấy giá trị "Doc Ref" trong Input file
                    string docTitle = wsInput.Cells[row, InputExcelData.DocTitle_Col].Text;

                    string verValue = wsInput.Cells[row, InputExcelData.Ver_Col].Text;
                    string revValue = wsInput.Cells[row, InputExcelData.REV_Col].Text;
                    string purposeValue = wsInput.Cells[row, InputExcelData.Purpose_Col].Text;
                    string statusValue = wsInput.Cells[row, InputExcelData.Status_Col].Text;
                    string transIdValue = wsInput.Cells[row, InputExcelData.TRANSID_Col].Text;
                    string dateValue_str = wsInput.Cells[row, InputExcelData.Date_Col].Text;
                    string timesCountValue = wsInput.Cells[row, InputExcelData.TimesCount_Col].Text;
                    string packageValue = wsInput.Cells[row, InputExcelData.Package_Col].Text;
                    string disciplineValue = wsInput.Cells[row, InputExcelData.Discipline_Col].Text;
                    string nativeFileTypeValue = wsInput.Cells[row, InputExcelData.NativeFileType_Col].Text;
                    string submittedFileTypeValue = wsInput.Cells[row, InputExcelData.SubmittedFileType_Col].Text;
                    string modelNameValue = wsInput.Cells[row, InputExcelData.ModelName_Col].Text;
                    string atlasCommentValue = wsInput.Cells[row, InputExcelData.AtlasComment_Col].Text;
                    string updateStatusValue = wsInput.Cells[row, InputExcelData.UpdateStatus_Col].Text;
                    string checkDocTitleValue = wsInput.Cells[row, InputExcelData.CheckDocTitle_Col].Text;

                    if (string.IsNullOrEmpty(docRef)) { continue; }

                    for (int masterRow = MasterExcelData_Drawing.Start_Row; masterRow <= totalRowsMaster; masterRow++)
                    {
                        string allianceNo = wsMaster.Cells[masterRow, MasterExcelData_Drawing.Alliance_Col].Text;   // Giá trị "Alliance No." trong Master file
                        if (allianceNo == docRef)
                        {
                            // DOC TITLE
                            if (checkDocTitleValue.StartsWith("Need to update"))
                            {
                                wsMaster.Cells[masterRow, docTitle_Master_Col].Value = docTitle;
                            }

                            // TIMES COUNT
                            wsMaster.Cells[masterRow, timesCount_Master_Col].Value = timesCountValue;
                            if (int.TryParse(timesCountValue, out int timesCountValue_int))
                            {
                                wsMaster.Cells[masterRow, timesCount_Master_Col].Value = timesCountValue_int;
                            }
                            // TRANS ID
                            wsMaster.Cells[masterRow, transid_Master_Col].Value = transIdValue;
                            // STATUS
                            wsMaster.Cells[masterRow, status_Master_Col].Value = statusValue;
                            // DATE
                            DateTime dateValue_ddmmyyy;
                            if (DateTime.TryParseExact(dateValue_str, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out dateValue_ddmmyyy))
                            {
                                wsMaster.Cells[masterRow, date_Master_Col].Value = dateValue_ddmmyyy;
                                wsMaster.Cells[masterRow, date_Master_Col].Style.Numberformat.Format = "dd/MM/yyyy"; // Định dạng ô theo ngày tháng năm
                            }
                            else
                            {
                                wsMaster.Cells[masterRow, date_Master_Col].Value = dateValue_str; // Giữ nguyên nếu không parse được
                            }
                            // ATLAS COMMENT
                            wsMaster.Cells[masterRow, atlasComment_Master_Col].Value = atlasCommentValue;

                            break;
                        }
                    }
                }

                // *** Đưa data vào InputRecordMaster
                TransferToInputRecordMaster(wsInput, wsInputRecordMaster,
                    docRef_InputRecordMaster_Col,
                    docTitle_InputRecordMaster_Col,
                    ver_InputRecordMaster_Col,
                    rev_InputRecordMaster_Col,
                    purpose_InputRecordMaster_Col,
                    status_InputRecordMaster_Col,
                    transid_InputRecordMaster_Col,
                    date_InputRecordMaster_Col,
                    timesCount_InputRecordMaster_Col,
                    package_InputRecordMaster_Col,
                    discipline_InputRecordMaster_Col,
                    redlineMarkup_InputRecordMaster_Col,
                    packageStampStatus_InputRecordMaster_Col,
                    date_Issue_InputRecordMaster_Col,
                    modelName_InputRecordMaster_Col,
                    nativeFileType_InputRecordMaster_Col,
                    submittedFileType_InputRecordMaster_Col,
                    atlasComment_InputRecordMaster_Col,
                    updateStatus_InputRecordMaster_Col,
                    checkDocTitle_InputRecordMaster_Col,
                    dateTime_Checked,
                    out isInput_Modified);

                // *** Phần cuối: Lưu các file
                packageMaster.Save();
                packageInputRecordMaster.Save();
                if (isInput_Modified) { packageInput.Save(); }

                MessageBox.Show(activeForm, "Data transfer is done!", caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            #endregion
        }
        //-------------------------------------------------------
        // 03. MODEL
        public static void Transfer_Model_Receive_ToMaster()
        {
            #region
            string caption = "Transfer Model_Receive to Master";
            string sheetName_InputRecorMaster = InputRecordMasterExcelData_Model_Receive.sheetName;

            if (!IsValidExcelFiles_ForTransfer_MainTab(activeForm, caption, sheetName_InputRecorMaster)) { return; }

            //--------------------------------------------
            // *** Backup Master và InputRecordMaster trước khi transfer data
            // * Backup Master File
            BackupFileToBackupFolder(txtFilePath_Master, Path.Combine("_backup", "backup_Master"));

            // * Backup InputRecordMaster File
            BackupFileToBackupFolder(txtFilePath_InputRecordMaster, Path.Combine("_backup", "backup_" + sheetName_InputRecorMaster));

            //--------------------------------------------
            // *** Transfer data
            bool isInput_Modified = false;

            // * Lấy giá trị ngày giờ trong filename của InputData file
            string dateTime_Checked = "";
            string fileName_InputData = Path.GetFileName(txtFilePath_InputData);
            string[] parts = fileName_InputData.Split('_');

            if (parts.Length >= 4)
            {
                dateTime_Checked = parts[parts.Length - 2] + "_" + parts[parts.Length - 1].Replace(".xlsx", "");
            }

            using (var packageMaster = new ExcelPackage(new FileInfo(txtFilePath_Master)))
            using (var packageInput = new ExcelPackage(new FileInfo(txtFilePath_InputData)))
            using (var packageInputRecordMaster = new ExcelPackage(new FileInfo(txtFilePath_InputRecordMaster)))
            {
                // * Ktra có Sheet cần tìm trong Master file ko rồi mới bắt đầu
                var wsMaster = GetWorksheetByName(packageMaster, MasterExcelData_Model.SheetName);
                if (wsMaster == null)
                {
                    MessageBox.Show(activeForm, "Sheet " + MasterExcelData_Model.SheetName + " not found in the file " + Path.GetFileName(txtFilePath_Master), "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                var wsInput = packageInput.Workbook.Worksheets[InputData_WorksheetIndex_Main];
                var wsInputRecordMaster = GetWorksheetByName(packageInputRecordMaster, sheetName_InputRecorMaster);

                // * Thông tin các cột của Master file
                int docTitle_Master_Col = MasterExcelData_Model.DocTitle_Col;
                int timesCount_Master_Col = MasterExcelData_Model_Receive.TimesCount_Col;
                int transid_Master_Col = MasterExcelData_Model_Receive.TRANSID_Col;
                int rev_Master_Col = MasterExcelData_Model_Receive.REV_Col;
                int ver_Master_Col = MasterExcelData_Model_Receive.Ver_Col;
                int purpose_Master_Col = MasterExcelData_Model_Receive.Purpose_Col;
                int date_Master_Col = MasterExcelData_Model_Receive.Date_Col;
                int status_Master_Col = MasterExcelData_Model_Receive.Status_Col;
                int atlasComment_Master_Col = MasterExcelData_Model_Receive.AtlasComment_Col;
                int fileType_Master_Col = MasterExcelData_Model_Receive.FileType_Col;

                // * Thông tin các cột của InputRecordMaster file
                int docRef_InputRecordMaster_Col = InputRecordMasterExcelData.DocRef_Col;
                int docTitle_InputRecordMaster_Col = InputRecordMasterExcelData.DocTitle_Col;
                int transid_InputRecordMaster_Col = InputRecordMasterExcelData.TRANSID_Col;

                int ver_InputRecordMaster_Col = InputRecordMasterExcelData_Model_Receive.Ver_Col;
                int rev_InputRecordMaster_Col = InputRecordMasterExcelData_Model_Receive.REV_Col;
                int purpose_InputRecordMaster_Col = InputRecordMasterExcelData_Model_Receive.Purpose_Col;
                int status_InputRecordMaster_Col = InputRecordMasterExcelData_Model_Receive.Status_Col;
                int date_InputRecordMaster_Col = InputRecordMasterExcelData_Model_Receive.Date_Col;
                int timesCount_InputRecordMaster_Col = InputRecordMasterExcelData_Model_Receive.TimesCount_Col;
                int package_InputRecordMaster_Col = InputRecordMasterExcelData_Model_Receive.Package_Col;
                int discipline_InputRecordMaster_Col = InputRecordMasterExcelData_Model_Receive.Discipline_Col;
                int redlineMarkup_InputRecordMaster_Col = -1;
                int packageStampStatus_InputRecordMaster_Col = -1;
                int date_Issue_InputRecordMaster_Col = -1;
                int modelName_InputRecordMaster_Col = -1;
                int nativeFileType_InputRecordMaster_Col = InputRecordMasterExcelData_Model_Receive.NativeFileType_Col;
                int submittedFileType_InputRecordMaster_Col = InputRecordMasterExcelData_Model_Receive.SubmittedFileType_Col;
                int atlasComment_InputRecordMaster_Col = InputRecordMasterExcelData_Model_Receive.AtlasComment_Col;
                int updateStatus_InputRecordMaster_Col = InputRecordMasterExcelData_Model_Receive.UpdateStatus_Col;
                int checkDocTitle_InputRecordMaster_Col = InputRecordMasterExcelData_Model_Receive.CheckDocTitle_Col;

                // * Lấy về tổng số dòng của các file
                int totalRowsInput = wsInput.Dimension.End.Row;
                int totalRowsMaster = wsMaster.Dimension.End.Row;

                // *** Đưa data vào Master
                for (int row = InputExcelData.Start_Row; row <= totalRowsInput; row++)
                {
                    // Chỉ lấy dòng có STATUS = OK
                    if (wsInput.Cells[row, InputExcelData.UpdateStatus_Col].Text != "OK")
                    {
                        continue; // Bỏ qua dòng Duplicate
                    }

                    string docRef = wsInput.Cells[row, InputExcelData.DocRef_Col].Text;     // Lấy giá trị "Doc Ref" trong Input file
                    string docTitle = wsInput.Cells[row, InputExcelData.DocTitle_Col].Text;

                    string verValue = wsInput.Cells[row, InputExcelData.Ver_Col].Text;
                    string revValue = wsInput.Cells[row, InputExcelData.REV_Col].Text;
                    string purposeValue = wsInput.Cells[row, InputExcelData.Purpose_Col].Text;
                    string statusValue = wsInput.Cells[row, InputExcelData.Status_Col].Text;
                    string transIdValue = wsInput.Cells[row, InputExcelData.TRANSID_Col].Text;
                    string dateValue_str = wsInput.Cells[row, InputExcelData.Date_Col].Text;
                    string timesCountValue = wsInput.Cells[row, InputExcelData.TimesCount_Col].Text;
                    string packageValue = wsInput.Cells[row, InputExcelData.Package_Col].Text;
                    string disciplineValue = wsInput.Cells[row, InputExcelData.Discipline_Col].Text;
                    string nativeFileTypeValue = wsInput.Cells[row, InputExcelData.NativeFileType_Col].Text;
                    string submittedFileTypeValue = wsInput.Cells[row, InputExcelData.SubmittedFileType_Col].Text;
                    string modelNameValue = wsInput.Cells[row, InputExcelData.ModelName_Col].Text;
                    string atlasCommentValue = wsInput.Cells[row, InputExcelData.AtlasComment_Col].Text;
                    string updateStatusValue = wsInput.Cells[row, InputExcelData.UpdateStatus_Col].Text;
                    string checkDocTitleValue = wsInput.Cells[row, InputExcelData.CheckDocTitle_Col].Text;

                    if (string.IsNullOrEmpty(docRef)) { continue; }

                    for (int masterRow = MasterExcelData_Model.Start_Row; masterRow <= totalRowsMaster; masterRow++)
                    {
                        string allianceNo = wsMaster.Cells[masterRow, MasterExcelData_Model.Alliance_Col].Text;   // Giá trị "Alliance No." trong Master file
                        if (allianceNo == docRef)
                        {
                            // DOC TITLE
                            if (checkDocTitleValue.StartsWith("Need to update"))
                            {
                                wsMaster.Cells[masterRow, docTitle_Master_Col].Value = docTitle;
                            }

                            // TIMES COUNT
                            wsMaster.Cells[masterRow, timesCount_Master_Col].Value = timesCountValue;
                            if (int.TryParse(timesCountValue, out int timesCountValue_int))
                            {
                                wsMaster.Cells[masterRow, timesCount_Master_Col].Value = timesCountValue_int;
                            }
                            // TRANS ID
                            wsMaster.Cells[masterRow, transid_Master_Col].Value = transIdValue;
                            // REV
                            wsMaster.Cells[masterRow, rev_Master_Col].Value = revValue;
                            if (int.TryParse(revValue, out int revValue_int))
                            {
                                wsMaster.Cells[masterRow, rev_Master_Col].Value = revValue_int;
                            }
                            // VER (ko có ở phần Submit)
                            wsMaster.Cells[masterRow, ver_Master_Col].Value = verValue;
                            if (int.TryParse(verValue, out int verValue_int))
                            {
                                wsMaster.Cells[masterRow, ver_Master_Col].Value = verValue_int;
                            }
                            // PURPOSE
                            wsMaster.Cells[masterRow, purpose_Master_Col].Value = purposeValue;
                            // DATE
                            DateTime dateValue_ddmmyyy;
                            if (DateTime.TryParseExact(dateValue_str, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out dateValue_ddmmyyy))
                            {
                                wsMaster.Cells[masterRow, date_Master_Col].Value = dateValue_ddmmyyy;
                                wsMaster.Cells[masterRow, date_Master_Col].Style.Numberformat.Format = "dd/MM/yyyy"; // Định dạng ô theo ngày tháng năm
                            }
                            else
                            {
                                wsMaster.Cells[masterRow, date_Master_Col].Value = dateValue_str; // Giữ nguyên nếu không parse được
                            }
                            // STATUS
                            wsMaster.Cells[masterRow, status_Master_Col].Value = statusValue;
                            // ATLAS COMMENT
                            wsMaster.Cells[masterRow, atlasComment_Master_Col].Value = atlasCommentValue;
                            // FILE TYPE
                            wsMaster.Cells[masterRow, fileType_Master_Col].Value = nativeFileTypeValue;

                            break;
                        }
                    }
                }

                // *** Đưa data vào InputRecordMaster
                TransferToInputRecordMaster(wsInput, wsInputRecordMaster,
                    docRef_InputRecordMaster_Col,
                    docTitle_InputRecordMaster_Col,
                    ver_InputRecordMaster_Col,
                    rev_InputRecordMaster_Col,
                    purpose_InputRecordMaster_Col,
                    status_InputRecordMaster_Col,
                    transid_InputRecordMaster_Col,
                    date_InputRecordMaster_Col,
                    timesCount_InputRecordMaster_Col,
                    package_InputRecordMaster_Col,
                    discipline_InputRecordMaster_Col,
                    redlineMarkup_InputRecordMaster_Col,
                    packageStampStatus_InputRecordMaster_Col,
                    date_Issue_InputRecordMaster_Col,
                    modelName_InputRecordMaster_Col,
                    nativeFileType_InputRecordMaster_Col,
                    submittedFileType_InputRecordMaster_Col,
                    atlasComment_InputRecordMaster_Col,
                    updateStatus_InputRecordMaster_Col,
                    checkDocTitle_InputRecordMaster_Col,
                    dateTime_Checked,
                    out isInput_Modified);

                // *** Phần cuối: Lưu các file
                packageMaster.Save();
                packageInputRecordMaster.Save();
                if (isInput_Modified) { packageInput.Save(); }

                MessageBox.Show(activeForm, "Data transfer is done!", caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            #endregion
        }
        public static void Transfer_Model_Submit_ToMaster()
        {
            #region
            string caption = "Transfer Model_Submit to Master";
            string sheetName_InputRecorMaster = InputRecordMasterExcelData_Model_Submit.sheetName;

            if (!IsValidExcelFiles_ForTransfer_MainTab(activeForm, caption, sheetName_InputRecorMaster)) { return; }

            //--------------------------------------------
            // *** Backup Master và InputRecordMaster trước khi transfer data
            // * Backup Master File
            BackupFileToBackupFolder(txtFilePath_Master, Path.Combine("_backup", "backup_Master"));

            // * Backup InputRecordMaster File
            BackupFileToBackupFolder(txtFilePath_InputRecordMaster, Path.Combine("_backup", "backup_" + sheetName_InputRecorMaster));

            //--------------------------------------------
            // *** Transfer data

            bool isInput_Modified = false;

            // * Lấy giá trị ngày giờ trong filename của InputData file
            string dateTime_Checked = "";
            string fileName_InputData = Path.GetFileName(txtFilePath_InputData);
            string[] parts = fileName_InputData.Split('_');

            if (parts.Length >= 4)
            {
                dateTime_Checked = parts[parts.Length - 2] + "_" + parts[parts.Length - 1].Replace(".xlsx", "");
            }

            using (var packageMaster = new ExcelPackage(new FileInfo(txtFilePath_Master)))
            using (var packageInput = new ExcelPackage(new FileInfo(txtFilePath_InputData)))
            using (var packageInputRecordMaster = new ExcelPackage(new FileInfo(txtFilePath_InputRecordMaster)))
            {
                // * Ktra có Sheet cần tìm trong Master file ko rồi mới bắt đầu
                var wsMaster = GetWorksheetByName(packageMaster, MasterExcelData_Model.SheetName);
                if (wsMaster == null)
                {
                    MessageBox.Show(activeForm, "Sheet " + MasterExcelData_Model.SheetName + " not found in the file " + Path.GetFileName(txtFilePath_Master), "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                var wsInput = packageInput.Workbook.Worksheets[InputData_WorksheetIndex_Main];
                var wsInputRecordMaster = GetWorksheetByName(packageInputRecordMaster, sheetName_InputRecorMaster);

                // * Thông tin các cột của Master file
                int timesCount_Master_Col = MasterExcelData_Model_Submission.TimesCount_Col;
                int transid_Master_Col = MasterExcelData_Model_Submission.TRANSID_Col;
                int rev_Master_Col = MasterExcelData_Model_Submission.REV_Col;
                int purpose_Master_Col = MasterExcelData_Model_Submission.Purpose_Col;
                int date_Master_Col = MasterExcelData_Model_Submission.Date_Col;
                int status_Master_Col = MasterExcelData_Model_Submission.Status_Col;
                int atlasComment_Master_Col = MasterExcelData_Model_Submission.AtlasComment_Col;
                int fileType_Master_Col = MasterExcelData_Model_Submission.FileType_Col;

                // * Thông tin các cột của InputRecordMaster file
                int docRef_InputRecordMaster_Col = InputRecordMasterExcelData.DocRef_Col;
                int docTitle_InputRecordMaster_Col = InputRecordMasterExcelData.DocTitle_Col;
                int transid_InputRecordMaster_Col = InputRecordMasterExcelData.TRANSID_Col;

                int ver_InputRecordMaster_Col = InputRecordMasterExcelData_Model_Submit.Ver_Col;
                int rev_InputRecordMaster_Col = InputRecordMasterExcelData_Model_Submit.REV_Col;
                int purpose_InputRecordMaster_Col = InputRecordMasterExcelData_Model_Submit.Purpose_Col;
                int status_InputRecordMaster_Col = InputRecordMasterExcelData_Model_Submit.Status_Col;
                int date_InputRecordMaster_Col = InputRecordMasterExcelData_Model_Submit.Date_Col;
                int timesCount_InputRecordMaster_Col = InputRecordMasterExcelData_Model_Submit.TimesCount_Col;
                int package_InputRecordMaster_Col = InputRecordMasterExcelData_Model_Submit.Package_Col;
                int discipline_InputRecordMaster_Col = InputRecordMasterExcelData_Model_Submit.Discipline_Col;
                int redlineMarkup_InputRecordMaster_Col = -1;
                int packageStampStatus_InputRecordMaster_Col = -1;
                int date_Issue_InputRecordMaster_Col = -1;
                int modelName_InputRecordMaster_Col = -1;
                int nativeFileType_InputRecordMaster_Col = InputRecordMasterExcelData_Model_Submit.NativeFileType_Col;
                int submittedFileType_InputRecordMaster_Col = InputRecordMasterExcelData_Model_Submit.SubmittedFileType_Col;
                int atlasComment_InputRecordMaster_Col = InputRecordMasterExcelData_Model_Submit.AtlasComment_Col;
                int updateStatus_InputRecordMaster_Col = InputRecordMasterExcelData_Model_Submit.UpdateStatus_Col;
                int checkDocTitle_InputRecordMaster_Col = InputRecordMasterExcelData_Model_Submit.CheckDocTitle_Col;

                // * Lấy về tổng số dòng của các file
                int totalRowsInput = wsInput.Dimension.End.Row;
                int totalRowsMaster = wsMaster.Dimension.End.Row;

                // *** Đưa data vào Master
                for (int row = InputExcelData.Start_Row; row <= totalRowsInput; row++)
                {
                    // Chỉ lấy dòng có STATUS = OK
                    if (wsInput.Cells[row, InputExcelData.UpdateStatus_Col].Text != "OK")
                    {
                        continue; // Bỏ qua dòng Duplicate
                    }

                    string docRef = wsInput.Cells[row, InputExcelData.DocRef_Col].Text;     // Lấy giá trị "Doc Ref" trong Input file
                    string docTitle = wsInput.Cells[row, InputExcelData.DocTitle_Col].Text;

                    string verValue = wsInput.Cells[row, InputExcelData.Ver_Col].Text;
                    string revValue = wsInput.Cells[row, InputExcelData.REV_Col].Text;
                    string purposeValue = wsInput.Cells[row, InputExcelData.Purpose_Col].Text;
                    string statusValue = wsInput.Cells[row, InputExcelData.Status_Col].Text;
                    string transIdValue = wsInput.Cells[row, InputExcelData.TRANSID_Col].Text;
                    string dateValue_str = wsInput.Cells[row, InputExcelData.Date_Col].Text;
                    string timesCountValue = wsInput.Cells[row, InputExcelData.TimesCount_Col].Text;
                    string packageValue = wsInput.Cells[row, InputExcelData.Package_Col].Text;
                    string disciplineValue = wsInput.Cells[row, InputExcelData.Discipline_Col].Text;
                    string nativeFileTypeValue = wsInput.Cells[row, InputExcelData.NativeFileType_Col].Text;
                    string submittedFileTypeValue = wsInput.Cells[row, InputExcelData.SubmittedFileType_Col].Text;
                    string modelNameValue = wsInput.Cells[row, InputExcelData.ModelName_Col].Text;
                    string atlasCommentValue = wsInput.Cells[row, InputExcelData.AtlasComment_Col].Text;
                    string updateStatusValue = wsInput.Cells[row, InputExcelData.UpdateStatus_Col].Text;
                    string checkDocTitleValue = wsInput.Cells[row, InputExcelData.CheckDocTitle_Col].Text;

                    if (string.IsNullOrEmpty(docRef)) { continue; }

                    for (int masterRow = MasterExcelData_Model.Start_Row; masterRow <= totalRowsMaster; masterRow++)
                    {
                        string allianceNo = wsMaster.Cells[masterRow, MasterExcelData_Model.Alliance_Col].Text;   // Giá trị "Alliance No." trong Master file
                        if (allianceNo == docRef)
                        {
                            // TIMES COUNT
                            wsMaster.Cells[masterRow, timesCount_Master_Col].Value = timesCountValue;
                            if (int.TryParse(timesCountValue, out int timesCountValue_int))
                            {
                                wsMaster.Cells[masterRow, timesCount_Master_Col].Value = timesCountValue_int;
                            }
                            // TRANS ID
                            wsMaster.Cells[masterRow, transid_Master_Col].Value = transIdValue;
                            // REV
                            wsMaster.Cells[masterRow, rev_Master_Col].Value = revValue;
                            if (int.TryParse(revValue, out int revValue_int))
                            {
                                wsMaster.Cells[masterRow, rev_Master_Col].Value = revValue_int;
                            }
                            // VER (ko có ở phần Submit)
                            //wsMaster.Cells[masterRow, ver_Master_Col].Value = verValue;
                            // PURPOSE
                            wsMaster.Cells[masterRow, purpose_Master_Col].Value = purposeValue;
                            // DATE
                            DateTime dateValue_ddmmyyy;
                            if (DateTime.TryParseExact(dateValue_str, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out dateValue_ddmmyyy))
                            {
                                wsMaster.Cells[masterRow, date_Master_Col].Value = dateValue_ddmmyyy;
                                wsMaster.Cells[masterRow, date_Master_Col].Style.Numberformat.Format = "dd/MM/yyyy"; // Định dạng ô theo ngày tháng năm
                            }
                            else
                            {
                                wsMaster.Cells[masterRow, date_Master_Col].Value = dateValue_str; // Giữ nguyên nếu không parse được
                            }
                            // STATUS
                            wsMaster.Cells[masterRow, status_Master_Col].Value = statusValue;
                            // ATLAS COMMENT
                            wsMaster.Cells[masterRow, atlasComment_Master_Col].Value = atlasCommentValue;
                            // FILE TYPE
                            wsMaster.Cells[masterRow, fileType_Master_Col].Value = submittedFileTypeValue;

                            break;
                        }
                    }
                }

                // *** Đưa data vào InputRecordMaster
                TransferToInputRecordMaster(wsInput, wsInputRecordMaster,
                    docRef_InputRecordMaster_Col,
                    docTitle_InputRecordMaster_Col,
                    ver_InputRecordMaster_Col,
                    rev_InputRecordMaster_Col,
                    purpose_InputRecordMaster_Col,
                    status_InputRecordMaster_Col,
                    transid_InputRecordMaster_Col,
                    date_InputRecordMaster_Col,
                    timesCount_InputRecordMaster_Col,
                    package_InputRecordMaster_Col,
                    discipline_InputRecordMaster_Col,
                    redlineMarkup_InputRecordMaster_Col,
                    packageStampStatus_InputRecordMaster_Col,
                    date_Issue_InputRecordMaster_Col,
                    modelName_InputRecordMaster_Col,
                    nativeFileType_InputRecordMaster_Col,
                    submittedFileType_InputRecordMaster_Col,
                    atlasComment_InputRecordMaster_Col,
                    updateStatus_InputRecordMaster_Col,
                    checkDocTitle_InputRecordMaster_Col,
                    dateTime_Checked,
                    out isInput_Modified);

                // *** Phần cuối: Lưu các file
                packageMaster.Save();
                packageInputRecordMaster.Save();
                if (isInput_Modified) { packageInput.Save(); }

                MessageBox.Show(activeForm, "Data transfer is done!", caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            #endregion
        }
        public static void Transfer_Model_Feedback_ToMaster()
        {
            #region
            string caption = "Transfer Model_Feedback to Master";
            string sheetName_InputRecorMaster = InputRecordMasterExcelData_Model_Feedback.sheetName;

            if (!IsValidExcelFiles_ForTransfer_MainTab(activeForm, caption, sheetName_InputRecorMaster)) { return; }

            //--------------------------------------------
            // *** Backup Master và InputRecordMaster trước khi transfer data
            // * Backup Master File
            BackupFileToBackupFolder(txtFilePath_Master, Path.Combine("_backup", "backup_Master"));

            // * Backup InputRecordMaster File
            BackupFileToBackupFolder(txtFilePath_InputRecordMaster, Path.Combine("_backup", "backup_" + sheetName_InputRecorMaster));

            //--------------------------------------------
            // *** Transfer data

            bool isInput_Modified = false;

            // * Lấy giá trị ngày giờ trong filename của InputData file
            string dateTime_Checked = "";
            string fileName_InputData = Path.GetFileName(txtFilePath_InputData);
            string[] parts = fileName_InputData.Split('_');

            if (parts.Length >= 4)
            {
                dateTime_Checked = parts[parts.Length - 2] + "_" + parts[parts.Length - 1].Replace(".xlsx", "");
            }

            using (var packageMaster = new ExcelPackage(new FileInfo(txtFilePath_Master)))
            using (var packageInput = new ExcelPackage(new FileInfo(txtFilePath_InputData)))
            using (var packageInputRecordMaster = new ExcelPackage(new FileInfo(txtFilePath_InputRecordMaster)))
            {
                // * Ktra có Sheet cần tìm trong Master file ko rồi mới bắt đầu
                var wsMaster = GetWorksheetByName(packageMaster, MasterExcelData_Model.SheetName);
                if (wsMaster == null)
                {
                    MessageBox.Show(activeForm, "Sheet " + MasterExcelData_Model.SheetName + " not found in the file " + Path.GetFileName(txtFilePath_Master), "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                var wsInput = packageInput.Workbook.Worksheets[InputData_WorksheetIndex_Main];
                var wsInputRecordMaster = GetWorksheetByName(packageInputRecordMaster, sheetName_InputRecorMaster);

                // * Thông tin các cột của Master file
                int timesCount_Master_Col = MasterExcelData_Model_Feedback.TimesCount_Col;
                int transid_Master_Col = MasterExcelData_Model_Feedback.TRANSID_Col;
                int status_Master_Col = MasterExcelData_Model_Feedback.Status_Col;
                int date_Master_Col = MasterExcelData_Model_Feedback.Date_Col;
                int atlasComment_Master_Col = MasterExcelData_Model_Feedback.AtlasComment_Col;

                // * Thông tin các cột của InputRecordMaster file
                int docRef_InputRecordMaster_Col = InputRecordMasterExcelData.DocRef_Col;
                int docTitle_InputRecordMaster_Col = InputRecordMasterExcelData.DocTitle_Col;
                int transid_InputRecordMaster_Col = InputRecordMasterExcelData.TRANSID_Col;

                int ver_InputRecordMaster_Col = InputRecordMasterExcelData_Model_Feedback.Ver_Col;
                int rev_InputRecordMaster_Col = InputRecordMasterExcelData_Model_Feedback.REV_Col;
                int purpose_InputRecordMaster_Col = InputRecordMasterExcelData_Model_Feedback.Purpose_Col;
                int status_InputRecordMaster_Col = InputRecordMasterExcelData_Model_Feedback.Status_Col;
                int date_InputRecordMaster_Col = InputRecordMasterExcelData_Model_Feedback.Date_Col;
                int timesCount_InputRecordMaster_Col = InputRecordMasterExcelData_Model_Feedback.TimesCount_Col;
                int package_InputRecordMaster_Col = InputRecordMasterExcelData_Model_Feedback.Package_Col;
                int discipline_InputRecordMaster_Col = InputRecordMasterExcelData_Model_Feedback.Discipline_Col;
                int redlineMarkup_InputRecordMaster_Col = -1;
                int packageStampStatus_InputRecordMaster_Col = -1;
                int date_Issue_InputRecordMaster_Col = -1;
                int modelName_InputRecordMaster_Col = InputRecordMasterExcelData_Model_Feedback.ModelName_Col;
                int nativeFileType_InputRecordMaster_Col = InputRecordMasterExcelData_Model_Feedback.NativeFileType_Col;
                int submittedFileType_InputRecordMaster_Col = InputRecordMasterExcelData_Model_Feedback.SubmittedFileType_Col;
                int atlasComment_InputRecordMaster_Col = InputRecordMasterExcelData_Model_Feedback.AtlasComment_Col;
                int updateStatus_InputRecordMaster_Col = InputRecordMasterExcelData_Model_Feedback.UpdateStatus_Col;
                int checkDocTitle_InputRecordMaster_Col = InputRecordMasterExcelData_Model_Feedback.CheckDocTitle_Col;

                // * Lấy về tổng số dòng của các file
                int totalRowsInput = wsInput.Dimension.End.Row;
                int totalRowsMaster = wsMaster.Dimension.End.Row;

                // *** Đưa data vào Master
                for (int row = InputExcelData.Start_Row; row <= totalRowsInput; row++)
                {
                    // Chỉ lấy dòng có STATUS = OK
                    if (wsInput.Cells[row, InputExcelData.UpdateStatus_Col].Text != "OK")
                    {
                        continue; // Bỏ qua dòng Duplicate
                    }

                    string docRef = wsInput.Cells[row, InputExcelData.DocRef_Col].Text;     // Lấy giá trị "Doc Ref" trong Input file
                    string docTitle = wsInput.Cells[row, InputExcelData.DocTitle_Col].Text;

                    string verValue = wsInput.Cells[row, InputExcelData.Ver_Col].Text;
                    string revValue = wsInput.Cells[row, InputExcelData.REV_Col].Text;
                    string purposeValue = wsInput.Cells[row, InputExcelData.Purpose_Col].Text;
                    string statusValue = wsInput.Cells[row, InputExcelData.Status_Col].Text;
                    string transIdValue = wsInput.Cells[row, InputExcelData.TRANSID_Col].Text;
                    string dateValue_str = wsInput.Cells[row, InputExcelData.Date_Col].Text;
                    string timesCountValue = wsInput.Cells[row, InputExcelData.TimesCount_Col].Text;
                    string packageValue = wsInput.Cells[row, InputExcelData.Package_Col].Text;
                    string disciplineValue = wsInput.Cells[row, InputExcelData.Discipline_Col].Text;
                    string nativeFileTypeValue = wsInput.Cells[row, InputExcelData.NativeFileType_Col].Text;
                    string submittedFileTypeValue = wsInput.Cells[row, InputExcelData.SubmittedFileType_Col].Text;
                    string modelNameValue = wsInput.Cells[row, InputExcelData.ModelName_Col].Text;
                    string atlasCommentValue = wsInput.Cells[row, InputExcelData.AtlasComment_Col].Text;
                    string updateStatusValue = wsInput.Cells[row, InputExcelData.UpdateStatus_Col].Text;
                    string checkDocTitleValue = wsInput.Cells[row, InputExcelData.CheckDocTitle_Col].Text;

                    if (string.IsNullOrEmpty(docRef)) { continue; }

                    for (int masterRow = MasterExcelData_Model.Start_Row; masterRow <= totalRowsMaster; masterRow++)
                    {
                        string allianceNo = wsMaster.Cells[masterRow, MasterExcelData_Model.Alliance_Col].Text;   // Giá trị "Alliance No." trong Master file
                        if (allianceNo == docRef)
                        {
                            // TIMES COUNT
                            wsMaster.Cells[masterRow, timesCount_Master_Col].Value = timesCountValue;
                            if (int.TryParse(timesCountValue, out int timesCountValue_int))
                            {
                                wsMaster.Cells[masterRow, timesCount_Master_Col].Value = timesCountValue_int;
                            }
                            // TRANS ID
                            wsMaster.Cells[masterRow, transid_Master_Col].Value = transIdValue;
                            // STATUS
                            wsMaster.Cells[masterRow, status_Master_Col].Value = statusValue;
                            // DATE
                            DateTime dateValue_ddmmyyy;
                            if (DateTime.TryParseExact(dateValue_str, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out dateValue_ddmmyyy))
                            {
                                wsMaster.Cells[masterRow, date_Master_Col].Value = dateValue_ddmmyyy;
                                wsMaster.Cells[masterRow, date_Master_Col].Style.Numberformat.Format = "dd/MM/yyyy"; // Định dạng ô theo ngày tháng năm
                            }
                            else
                            {
                                wsMaster.Cells[masterRow, date_Master_Col].Value = dateValue_str; // Giữ nguyên nếu không parse được
                            }
                            // ATLAS COMMENT
                            wsMaster.Cells[masterRow, atlasComment_Master_Col].Value = atlasCommentValue;

                            break;
                        }
                    }
                }

                // *** Đưa data vào InputRecordMaster
                TransferToInputRecordMaster(wsInput, wsInputRecordMaster,
                    docRef_InputRecordMaster_Col,
                    docTitle_InputRecordMaster_Col,
                    ver_InputRecordMaster_Col,
                    rev_InputRecordMaster_Col,
                    purpose_InputRecordMaster_Col,
                    status_InputRecordMaster_Col,
                    transid_InputRecordMaster_Col,
                    date_InputRecordMaster_Col,
                    timesCount_InputRecordMaster_Col,
                    package_InputRecordMaster_Col,
                    discipline_InputRecordMaster_Col,
                    redlineMarkup_InputRecordMaster_Col,
                    packageStampStatus_InputRecordMaster_Col,
                    date_Issue_InputRecordMaster_Col,
                    modelName_InputRecordMaster_Col,
                    nativeFileType_InputRecordMaster_Col,
                    submittedFileType_InputRecordMaster_Col,
                    atlasComment_InputRecordMaster_Col,
                    updateStatus_InputRecordMaster_Col,
                    checkDocTitle_InputRecordMaster_Col,
                    dateTime_Checked,
                    out isInput_Modified);

                // *** Phần cuối: Lưu các file
                packageMaster.Save();
                packageInputRecordMaster.Save();
                if (isInput_Modified) { packageInput.Save(); }

                MessageBox.Show(activeForm, "Data transfer is done!", caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            #endregion
        }

        //============================================================================
        // ======TRANSFER DATA - TIMESHEET TAB
        public static void Transfer_Drawing_TimeSheet()
        {
            #region
            string caption = "Transfer Drawing_TimeSheet to Master";
            
            string sheetName_TS_Input = "Drawing List";
            string sheetName_TS_InputRecorMaster = "Drawings";

            if (!IsValidExcelFiles_ForTransfer_TSTab(activeForm, caption)) { return; }

            //--------------------------------------------
            // *** Backup Master và InputRecordMaster trước khi transfer data
            // * Backup Master File
            BackupFileToBackupFolder(txtFilePath_Master, Path.Combine("_backup", "backup_Master"));

            // * Backup InputRecordMaster File
            BackupFileToBackupFolder(txtFilePath_TS_InputRecordMaster, Path.Combine("_backup", "backup_TS_" + sheetName_TS_InputRecorMaster));

            // Xác định vị trí các cột dựa vào max Uniquecode hoặc đặt mặc định
            int totalHour_Input_Col = GetColumnIndexByUniqueCode(txtFilePath_TS_InputData, sheetName_TS_Input, "total#");
            int wpr_InputRecordMaster_StartCol = GetColumnIndexByUniqueCode(txtFilePath_TS_InputRecordMaster, sheetName_TS_InputRecorMaster, "wpr#");

            //--------------------------------------------
            // *** Transfer data
            using (var packageMaster = new ExcelPackage(new FileInfo(txtFilePath_Master)))
            using (var packageInputRecordMaster = new ExcelPackage(new FileInfo(txtFilePath_TS_InputRecordMaster)))
            using (var packageInput = new ExcelPackage(new FileInfo(txtFilePath_TS_InputData)))
            {
                // * Ktra có Sheet cần tìm trong Master file ko rồi mới bắt đầu
                var wsMaster = GetWorksheetByName(packageMaster, MasterExcelData_Drawing.SheetName);
                if (wsMaster == null)
                {
                    MessageBox.Show(activeForm, "Sheet " + MasterExcelData_Drawing.SheetName + " not found in the file " + Path.GetFileName(txtFilePath_Master), "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                var wsInput = GetWorksheetByName(packageInput, sheetName_TS_Input);
                if (wsInput == null)
                {
                    MessageBox.Show(activeForm, "Không tìm thấy sheet " + sheetName_TS_Input + " trong Input file", "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                var wsInputRecordMaster = GetWorksheetByName(packageInputRecordMaster, sheetName_TS_InputRecorMaster);

                // 1. Tìm cột trống trong file Master (từ cột S trở đi, hàng 12)
                int startColumn_Master = MasterExcelData_Drawing.WPR_Col; // Cột bắt đầu ghi WPR
                int emptyColumn_Master = -1;
                for (int col = startColumn_Master; col <= 1000; col++)
                {
                    if (wsMaster.Cells[MasterExcelData_Drawing.Start_Row - 2, col].Value == null)
                    {
                        emptyColumn_Master = col;
                        break;
                    }
                }

                // 2. Điền dữ liệu từ file Input vào file Master
                // Lấy giá trị từ Input B1 và D2
                string reportNo_Input_Value = wsInput.Cells[1, 2].Text;
                string dateReport_Input_Value = wsInput.Cells[2, 4].Text;

                wsMaster.Cells[MasterExcelData_Drawing.Start_Row - 2, emptyColumn_Master].Value = reportNo_Input_Value;

                // Chuyển đổi định dạng ngày tháng
                DateTime parsedDate;
                if (DateTime.TryParseExact(dateReport_Input_Value, "ddd, MMM dd, yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out parsedDate))
                {
                    wsMaster.Cells[MasterExcelData_Drawing.Start_Row - 1, emptyColumn_Master].Value = parsedDate;
                    wsMaster.Cells[MasterExcelData_Drawing.Start_Row - 1, emptyColumn_Master].Style.Numberformat.Format = "dd/MM/yyyy";
                }
                else
                {
                    wsMaster.Cells[MasterExcelData_Drawing.Start_Row - 1, emptyColumn_Master].Value = dateReport_Input_Value;
                }

                // 3. Lưu trước danh sách Alliance từ file Input vào Dictionary
                //Key = Alliance (cột C) / Value = Giá trị cột Total Hour (nếu > 0)
                Dictionary<string, double> alliance_InputData = new Dictionary<string, double>();
                int lastRowInput = wsInput.Dimension.Rows;

                for (int rowInput = 5; rowInput <= lastRowInput; rowInput++)
                {
                    string alliance_Input = wsInput.Cells[rowInput, 3].Text; // Cột Alliance file Input
                    if (double.TryParse(wsInput.Cells[rowInput, totalHour_Input_Col].Text, out double value_Input) && value_Input > 0) // Cột Total file Input
                    {
                        if (!alliance_InputData.ContainsKey(alliance_Input))
                        {
                            alliance_InputData[alliance_Input] = value_Input;
                        }
                    }
                }

                // 4. Duyệt file Master, kiểm tra Alliance trong Dictionary và cập nhật giá trị
                int lastRowMaster = wsMaster.Dimension.Rows;

                for (int rowMaster = MasterExcelData_Drawing.Start_Row; rowMaster <= lastRowMaster; rowMaster++)
                {
                    string alliance_Master = wsMaster.Cells[rowMaster, MasterExcelData_Drawing.Alliance_Col].Text; // Cột Alliance file Master
                    if (alliance_InputData.TryGetValue(alliance_Master, out double matchedValue))
                    {
                        wsMaster.Cells[rowMaster, emptyColumn_Master].Value = matchedValue;
                    }
                }

                // 5. Đưa data vào InputRecordMaster
                TransferToInputRecordMaster_TS(wsInput, wsInputRecordMaster,
                    totalHour_Input_Col,
                    wpr_InputRecordMaster_StartCol,
                    reportNo_Input_Value,
                    dateReport_Input_Value);

                // 6. Lưu file Master
                packageMaster.Save();
                packageInputRecordMaster.Save();

                MessageBox.Show(activeForm, "Data transfer is done!", caption, 
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            #endregion
        }

        public static void Transfer_Model_TimeSheet()
        {
            #region
            string caption = "Transfer Model_TimeSheet to Master";
            
            string sheetName_TS_Input = "Model List";
            string sheetName_TS_InputRecorMaster = "Model";

            if (!IsValidExcelFiles_ForTransfer_TSTab(activeForm, caption)) { return; }

            //--------------------------------------------
            // *** Backup Master và InputRecordMaster trước khi transfer data
            // * Backup Master File
            BackupFileToBackupFolder(txtFilePath_Master, Path.Combine("_backup", "backup_Master"));

            // * Backup InputRecordMaster File
            BackupFileToBackupFolder(txtFilePath_TS_InputRecordMaster, Path.Combine("_backup", "backup_TS_" + sheetName_TS_InputRecorMaster));

            // Xác định vị trí các cột dựa vào max Uniquecode hoặc đặt mặc định
            int totalHour_Input_Col = GetColumnIndexByUniqueCode(txtFilePath_TS_InputData, sheetName_TS_Input, "total#");
            int wpr_InputRecordMaster_StartCol = GetColumnIndexByUniqueCode(txtFilePath_TS_InputRecordMaster, sheetName_TS_InputRecorMaster, "wpr#");

            //--------------------------------------------
            // *** Transfer data
            using (var packageMaster = new ExcelPackage(new FileInfo(txtFilePath_Master)))
            using (var packageInputRecordMaster = new ExcelPackage(new FileInfo(txtFilePath_TS_InputRecordMaster)))
            using (var packageInput = new ExcelPackage(new FileInfo(txtFilePath_TS_InputData)))
            {
                // * Ktra có Sheet cần tìm trong Master file ko rồi mới bắt đầu
                var wsMaster = GetWorksheetByName(packageMaster, MasterExcelData_Model.SheetName);
                if (wsMaster == null)
                {
                    MessageBox.Show(activeForm, "Sheet " + MasterExcelData_Model.SheetName + " not found in the file " + Path.GetFileName(txtFilePath_Master), "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                var wsInput = GetWorksheetByName(packageInput, sheetName_TS_Input);
                if (wsInput == null)
                {
                    MessageBox.Show(activeForm, "Không tìm thấy sheet " + sheetName_TS_Input + " trong Input file", "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                var wsInputRecordMaster = GetWorksheetByName(packageInputRecordMaster, sheetName_TS_InputRecorMaster);

                // 1. Tìm cột trống trong file Master (từ cột S trở đi, hàng 12)
                int startColumn_Master = MasterExcelData_Model.WPR_Col; // Cột bắt đầu ghi WPR
                int emptyColumn_Master = -1;
                for (int col = startColumn_Master; col <= 1000; col++)
                {
                    if (wsMaster.Cells[MasterExcelData_Model.Start_Row - 2, col].Value == null)
                    {
                        emptyColumn_Master = col;
                        break;
                    }
                }

                // 2. Điền dữ liệu từ file Input vào file Master
                // Lấy giá trị từ Input B1 và D2
                string reportNo_Input_Value = wsInput.Cells[1, 2].Text;
                string dateReport_Input_Value = wsInput.Cells[2, 4].Text;

                wsMaster.Cells[MasterExcelData_Model.Start_Row - 2, emptyColumn_Master].Value = reportNo_Input_Value;

                // Chuyển đổi định dạng ngày tháng
                DateTime parsedDate;
                if (DateTime.TryParseExact(dateReport_Input_Value, "ddd, MMM dd, yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out parsedDate))
                {
                    wsMaster.Cells[MasterExcelData_Model.Start_Row - 1, emptyColumn_Master].Value = parsedDate;
                    wsMaster.Cells[MasterExcelData_Model.Start_Row - 1, emptyColumn_Master].Style.Numberformat.Format = "dd/MM/yyyy";
                }
                else
                {
                    wsMaster.Cells[MasterExcelData_Model.Start_Row - 1, emptyColumn_Master].Value = dateReport_Input_Value;
                }

                // 3. Lưu trước danh sách Alliance từ file Input vào Dictionary
                //Key = Alliance (cột C) / Value = Giá trị cột Total Hour (nếu > 0)
                Dictionary<string, double> alliance_InputData = new Dictionary<string, double>();
                int lastRowInput = wsInput.Dimension.Rows;

                for (int rowInput = 5; rowInput <= lastRowInput; rowInput++)
                {
                    string alliance_Input = wsInput.Cells[rowInput, 3].Text; // Cột Alliance file Input
                    if (double.TryParse(wsInput.Cells[rowInput, totalHour_Input_Col].Text, out double value_Input) && value_Input > 0) // Cột Total file Input
                    {
                        if (!alliance_InputData.ContainsKey(alliance_Input))
                        {
                            alliance_InputData[alliance_Input] = value_Input;
                        }
                    }
                }

                // 4. Duyệt file Master, kiểm tra Alliance trong Dictionary và cập nhật giá trị
                int lastRowMaster = wsMaster.Dimension.Rows;

                for (int rowMaster = MasterExcelData_Model.Start_Row; rowMaster <= lastRowMaster; rowMaster++)
                {
                    string alliance_Master = wsMaster.Cells[rowMaster, MasterExcelData_Model.Alliance_Col].Text; // Cột Alliance file Master
                    if (alliance_InputData.TryGetValue(alliance_Master, out double matchedValue))
                    {
                        wsMaster.Cells[rowMaster, emptyColumn_Master].Value = matchedValue;
                    }
                }

                // 5. Đưa data vào InputRecordMaster
                TransferToInputRecordMaster_TS(wsInput, wsInputRecordMaster,
                    totalHour_Input_Col,
                    wpr_InputRecordMaster_StartCol,
                    reportNo_Input_Value,
                    dateReport_Input_Value);

                // 6. Lưu file Master
                packageMaster.Save();
                packageInputRecordMaster.Save();

                MessageBox.Show(activeForm, "Data transfer is done!", caption, 
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            #endregion
        }

        //============================================================================
        // ======GET FILE NAME - GET FILES IN FOLDER TAB
        public static void GetFileName_FromInputFolder()
        {
            #region
            string sourceInputTemplateFolder = @"\\SRVPRD4\Structure\ATLAS TOOLS - STR\Excel Tracking";
            string templateInputFileName = "Input Data_Template.xlsx";
            string sourceInputTemplateFile = Path.Combine(sourceInputTemplateFolder, templateInputFileName);
            
            string inputDataFolder = txtFilePath_GetFile_InputDataFolder;
            string outputFolder = txtFilePath_GetFile_OutputDataFolder;
            string excelInputFilePath = Path.Combine(outputFolder, templateInputFileName);

            int inputData_Start_Row = 11;
            int inputData_DocRef_Col = 2;
            int inputData_DocTitle_Col = 3;
            int inputData_Rev_Col = 5;
            int inputData_Purpose_Col = 6;
            int inputData_PTANo_Col = 16;
            int inputData_Date_FromCAD_Col = 17;
            int inputData_NativeFileType = 13;

            int inputData_Check_Rev_InCADFile_Col = 24;
            int inputData_Rev_MPDT_Col = 25;
            int inputData_Check_Rev_WithMPDT = 26;
            int inputData_CheckDataExistingInMaster_Col = 27;

            int mainColum = inputData_DocRef_Col;
            if (isInputAsPTANo)
            {
                mainColum = inputData_PTANo_Col;
            }

            if (!Directory.Exists(inputDataFolder))
            {
                MessageBox.Show(activeForm, "Input data folder does not exist", "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (!File.Exists(sourceInputTemplateFile))
            {
                MessageBox.Show(activeForm, "The Excel template file " + Path.GetFileName(excelInputFilePath) + " does not exist!", "Warning",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (!Directory.Exists(outputFolder))
            {
                MessageBox.Show(activeForm, "Output folder does not exist!", "Warning",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Sao chép input template file từ nguồn đến đích
            File.Copy(sourceInputTemplateFile, excelInputFilePath, true);

            // Lấy danh sách tất cả tệp trong thư mục
            string[] files = Directory.GetFiles(inputDataFolder);

            FileInfo existingFile = new FileInfo(excelInputFilePath);
            using (var packageMaster = new ExcelPackage(new FileInfo(txtFilePath_Master)))
            using (var packageInput = new ExcelPackage(existingFile))
            {
                var wsInput = packageInput.Workbook.Worksheets[2]; // Lấy sheet 2
                if (wsInput == null)
                {
                    MessageBox.Show(activeForm, "Không thấy Worksheet", "Warning",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Ghi dữ liệu vào Excel
                for (int i = 0; i < files.Length; i++)
                {
                    int row = inputData_Start_Row + i;
                    wsInput.Cells[row, mainColum].Value = Path.GetFileNameWithoutExtension(files[i]); // Tên tệp
                    wsInput.Cells[row, inputData_NativeFileType].Value = Path.GetExtension(files[i]).Substring(1); // Định dạng file extension
                }

                // Thêm tiêu đề cho các cột từ 24 - 27
                wsInput.Cells[inputData_Start_Row - 1, inputData_Check_Rev_InCADFile_Col].Value = "Check Rev in CAD file";
                wsInput.Cells[inputData_Start_Row - 1, inputData_Check_Rev_InCADFile_Col].Style.WrapText = true;
                wsInput.Cells[inputData_Start_Row - 1, inputData_Rev_MPDT_Col].Value = "Rev MPDT";
                wsInput.Cells[inputData_Start_Row - 1, inputData_Rev_MPDT_Col].Style.WrapText = true;
                wsInput.Cells[inputData_Start_Row - 1, inputData_Check_Rev_WithMPDT].Value = "Check Rev w MPDT";
                wsInput.Cells[inputData_Start_Row - 1, inputData_Check_Rev_WithMPDT].Style.WrapText = true;
                wsInput.Cells[inputData_Start_Row - 1, inputData_CheckDataExistingInMaster_Col].Value = "Check DocRef/PTANo in MPDT";
                wsInput.Cells[inputData_Start_Row - 1, inputData_CheckDataExistingInMaster_Col].Style.WrapText = true;

                // * Ktra có Sheet cần tìm trong Master file ko rồi mới bắt đầu
                ExcelWorksheet wsMaster = null;
                if (!isInputAsModel)
                {
                    wsMaster = GetWorksheetByName(packageMaster, MasterExcelData_Drawing.SheetName);
                    if (wsMaster == null)
                    {
                        MessageBox.Show(activeForm, "Sheet " + MasterExcelData_Drawing.SheetName + " not found in the file " + Path.GetFileName(txtFilePath_Master), "Error",
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
                else
                {
                    wsMaster = GetWorksheetByName(packageMaster, MasterExcelData_Model.SheetName);
                    if (wsMaster == null)
                    {
                        MessageBox.Show(activeForm, "Sheet " + MasterExcelData_Drawing.SheetName + " not found in the file " + Path.GetFileName(txtFilePath_Master), "Error",
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }

                // * Thông tin các cột của Master file
                int totalRowsMaster = wsMaster.Dimension.End.Row;
                int startRow_Master = MasterExcelData_Drawing.Start_Row;
                
                int alliance_Master_Col = MasterExcelData_Drawing.Alliance_Col;
                int ptaNo_Master_Col = alliance_Master_Col + 1;
                int docTitle_Master_Col = MasterExcelData_Drawing.DocTitle_Col;
                int rev_First_Master_Col = MasterExcelData_Drawing_FirstReceive.REV_Col;

                if (isInputAsModel)
                {
                    startRow_Master = MasterExcelData_Model.Start_Row;
                    alliance_Master_Col = MasterExcelData_Model.Alliance_Col;
                    ptaNo_Master_Col = alliance_Master_Col + 1;
                    docTitle_Master_Col = MasterExcelData_Model.DocTitle_Col;
                    rev_First_Master_Col = MasterExcelData_Model_Receive.REV_Col;
                }

                if (isInputAsPTANo)
                {
                    for (int i = 0; i <= files.Length; i++)
                    {
                        int row = inputData_Start_Row + i;
                        string ptaNo_Input = wsInput.Cells[row, inputData_PTANo_Col].Text;    // Lấy giá trị PTANo từ Input file
                        if (string.IsNullOrEmpty(ptaNo_Input)) { continue; }

                        // Đưa các giá trị Rev - Purpose - Datetime từ CAD file vào inputdata
                        if (isGetCADInfo)
                        {
                            if (i < files.Length)
                            {
                                CadInfoExtractor.GetCadInfo(files[i],
                                    out string maxRevValue, out string maxRevDate, out string maxRevAmendment, out string status_CheckRev_InCAD);

                                if (!string.IsNullOrEmpty(maxRevValue))
                                {
                                    // REV
                                    wsInput.Cells[row, inputData_Rev_Col].Value = maxRevValue;
                                    if (int.TryParse(maxRevValue, out int maxRevValue_int))
                                    {
                                        wsInput.Cells[row, inputData_Rev_Col].Value = maxRevValue_int;
                                    }

                                    // PURPOSE
                                    wsInput.Cells[row, inputData_Purpose_Col].Value = maxRevAmendment;

                                    // DateTime
                                    string[] dateFormats =
                                        {
                                                    "d/M/yyyy", "d/M/yy", "dd/MM/yyyy", "dd/MM/yy",
                                                    "d/MM/yyyy", "dd/M/yyyy", "d/MM/yy", "dd/M/yy"
                                                };
                                    DateTime dateValue;
                                    if (DateTime.TryParseExact(maxRevDate, dateFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out dateValue))
                                    {
                                        wsInput.Cells[row, inputData_Date_FromCAD_Col].Value = dateValue;
                                        wsInput.Cells[row, inputData_Date_FromCAD_Col].Style.Numberformat.Format = "dd/MM/yyyy"; // Định dạng ô theo ngày tháng năm
                                    }
                                    else
                                    {
                                        wsInput.Cells[row, inputData_Date_FromCAD_Col].Value = maxRevDate; // Giữ nguyên nếu không parse được
                                    }

                                    if (!string.IsNullOrEmpty(status_CheckRev_InCAD))
                                    {
                                        wsInput.Cells[row, inputData_Check_Rev_InCADFile_Col].Value = status_CheckRev_InCAD;
                                    }
                                }
                            }
                        }

                        bool isFound = false;
                        for (int masterRow = startRow_Master; masterRow <= totalRowsMaster; masterRow++)
                        {
                            string allianceNo_Master = wsMaster.Cells[masterRow, alliance_Master_Col].Text;
                            string ptaNo_Master = wsMaster.Cells[masterRow, ptaNo_Master_Col].Text;
                            string docTitle_Master = wsMaster.Cells[masterRow, docTitle_Master_Col].Text;
                            string rev_Drawing_First_Master = wsMaster.Cells[masterRow, rev_First_Master_Col].Text;

                            if (ptaNo_Input == ptaNo_Master)
                            {
                                isFound = true;

                                wsInput.Cells[row, inputData_DocRef_Col].Value = allianceNo_Master;
                                wsInput.Cells[row, inputData_DocTitle_Col].Value = docTitle_Master;
                                wsInput.Cells[row, inputData_Rev_MPDT_Col].Value = rev_Drawing_First_Master;
                                if (int.TryParse(rev_Drawing_First_Master, out int revValue_int))
                                {
                                    wsInput.Cells[row, inputData_Rev_MPDT_Col].Value = revValue_int;
                                }

                                if (rev_Drawing_First_Master != wsInput.Cells[row, inputData_Rev_Col].Text)
                                {
                                    wsInput.Cells[row, inputData_Check_Rev_WithMPDT].Value = "Rev Not Match";
                                }
                                
                                break;
                            }
                        }
                        if (!isFound)
                        {
                            wsInput.Cells[row, inputData_CheckDataExistingInMaster_Col].Value = "Not found in MPDT";
                            wsInput.Cells[row, inputData_CheckDataExistingInMaster_Col].Style.Font.Color.SetColor(Color.Red);
                        }
                    }
                }
                else
                {
                    for (int i = 0; i <= files.Length; i++)
                    {
                        int row = inputData_Start_Row + i;
                        string docRef_Input = wsInput.Cells[row, inputData_DocRef_Col].Text;    // Lấy giá trị DocRef từ Input file
                        if (string.IsNullOrEmpty(docRef_Input)) { continue; }

                        // Đưa các giá trị Rev - Purpose - Datetime từ CAD file vào inputdata
                        if (isGetCADInfo)
                        {
                            CadInfoExtractor.GetCadInfo(files[i],
                                    out string maxRevValue, out string maxRevDate, out string maxRevAmendment, out string status_CheckRev_InCAD);

                            if (!string.IsNullOrEmpty(maxRevValue))
                            {
                                // REV
                                wsInput.Cells[row, inputData_Rev_Col].Value = maxRevValue;
                                if (int.TryParse(maxRevValue, out int maxRevValue_int))
                                {
                                    wsInput.Cells[row, inputData_Rev_Col].Value = maxRevValue_int;
                                }

                                // PURPOSE
                                wsInput.Cells[row, inputData_Purpose_Col].Value = maxRevAmendment;

                                // DateTime
                                string[] dateFormats =
                                        {
                                                    "d/M/yyyy", "d/M/yy", "dd/MM/yyyy", "dd/MM/yy",
                                                    "d/MM/yyyy", "dd/M/yyyy", "d/MM/yy", "dd/M/yy"
                                                };
                                DateTime dateValue;
                                if (DateTime.TryParseExact(maxRevDate, dateFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out dateValue))
                                {
                                    wsInput.Cells[row, inputData_Date_FromCAD_Col].Value = dateValue;
                                    wsInput.Cells[row, inputData_Date_FromCAD_Col].Style.Numberformat.Format = "dd/MM/yyyy"; // Định dạng ô theo ngày tháng năm
                                }
                                else
                                {
                                    wsInput.Cells[row, inputData_Date_FromCAD_Col].Value = maxRevDate; // Giữ nguyên nếu không parse được
                                }

                                if (!string.IsNullOrEmpty(status_CheckRev_InCAD))
                                {
                                    wsInput.Cells[row, inputData_Check_Rev_InCADFile_Col].Value = status_CheckRev_InCAD;
                                }
                            }
                        }

                        bool isFound = false;
                        for (int masterRow = startRow_Master; masterRow <= totalRowsMaster; masterRow++)
                        {
                            string allianceNo_Master = wsMaster.Cells[masterRow, alliance_Master_Col].Text;
                            string ptaNo_Master = wsMaster.Cells[masterRow, ptaNo_Master_Col].Text;
                            string docTitle_Master = wsMaster.Cells[masterRow, docTitle_Master_Col].Text;
                            string rev_Drawing_First_Master = wsMaster.Cells[masterRow, rev_First_Master_Col].Text;

                            if (docRef_Input == allianceNo_Master)
                            {
                                wsInput.Cells[row, inputData_PTANo_Col].Value = ptaNo_Master;
                                wsInput.Cells[row, inputData_DocTitle_Col].Value = docTitle_Master;
                                wsInput.Cells[row, inputData_Rev_MPDT_Col].Value = rev_Drawing_First_Master;
                                if (int.TryParse(rev_Drawing_First_Master, out int revValue_int))
                                {
                                    wsInput.Cells[row, inputData_Rev_MPDT_Col].Value = revValue_int;
                                }

                                if (rev_Drawing_First_Master != wsInput.Cells[row, inputData_Rev_Col].Text)
                                {
                                    wsInput.Cells[row, inputData_Check_Rev_WithMPDT].Value = "Rev Not Match";
                                }

                                isFound = true;
                                break;
                            }
                        }
                        if (!isFound)
                        {
                            wsInput.Cells[row, inputData_CheckDataExistingInMaster_Col].Value = "Not found in MPDT";
                            wsInput.Cells[row, inputData_CheckDataExistingInMaster_Col].Style.Font.Color.SetColor(Color.Red);
                        }
                    }
                }

                wsInput.Select();
                packageInput.Save();
            }
            // Mở thư mục đích
            System.Diagnostics.Process.Start("explorer.exe", outputFolder);
            #endregion
        }

        //============================================================================
        // ======TRANSFER DATA - OUTPUT TAB
        public static void ExportData_OutputForm()
        {
            #region
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
                openFileDialog.Title = "Select Input Data Excel File";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string excelInputFilePath = openFileDialog.FileName;

                    if (!File.Exists(excelInputFilePath))
                    {
                        MessageBox.Show(activeForm, "The Excel file " + Path.GetFileName(excelInputFilePath) + " does not exist!", "Warning",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    FileInfo existingFile = new FileInfo(excelInputFilePath);
                    using (var packageInput = new ExcelPackage(existingFile))
                    {
                        var wsInput_InputData = packageInput.Workbook.Worksheets[1]; // Lấy sheet 1
                        var wsInput_OutputData = packageInput.Workbook.Worksheets[3]; // Lấy sheet 3
                        if (wsInput_InputData == null || wsInput_OutputData == null)
                        {
                            MessageBox.Show(activeForm, "Không thấy Worksheet", "Warning",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }

                        int inputData_StartRow = 11;
                        int inputData_No_Col = 1;
                        int inputData_DocRef_Col = 2;
                        int inputData_DocTitle_Col = 3;
                        int inputData_Purpose_Col = 6;
                        int inputData_Status_Col = 7;
                        int inputData_PTANo_Col = 18;

                        int outputData_StartRow = 3;
                        int outputData_No_Col = 1;
                        int outputData_DocRef_Col = 2;
                        int outputData_PTANo_Col = 3;
                        int outputData_DocTitle_Col = 4;
                        int outputData_Purpose_Col = 5;
                        int outputData_Status_Col = 6;

                        // Phần transfer data từ sheet1 sang sheet3
                        int outputRow = outputData_StartRow;

                        // Xác định số dòng có dữ liệu trong sheet1
                        int inputRowCount = wsInput_InputData.Dimension.End.Row;

                        // Duyệt qua từng dòng của sheet1 bắt đầu từ inputData_StartRow
                        for (int inputRow = inputData_StartRow; inputRow <= inputRowCount; inputRow++)
                        {
                            // Kiểm tra cột DocRef có giá trị không
                            string docRef = wsInput_InputData.Cells[inputRow, inputData_DocRef_Col].Text;

                            // Nếu DocRef có giá trị thì transfer data sang sheet3
                            if (!string.IsNullOrWhiteSpace(docRef))
                            {
                                // Lấy các giá trị từ sheet1
                                string number = wsInput_InputData.Cells[inputRow, inputData_No_Col].Text;
                                string docTitle = wsInput_InputData.Cells[inputRow, inputData_DocTitle_Col].Text;
                                string purpose = wsInput_InputData.Cells[inputRow, inputData_Purpose_Col].Text;
                                string status = wsInput_InputData.Cells[inputRow, inputData_Status_Col].Text;
                                string ptaNo = wsInput_InputData.Cells[inputRow, inputData_PTANo_Col].Text;

                                // Điền vào sheet3 với các cột tương ứng
                                wsInput_OutputData.Cells[outputRow, outputData_No_Col].Value = number;
                                if (int.TryParse(number, out int numbere_int))
                                {
                                    wsInput_OutputData.Cells[outputRow, outputData_No_Col].Value = numbere_int;
                                }

                                wsInput_OutputData.Cells[outputRow, outputData_DocRef_Col].Value = docRef;
                                wsInput_OutputData.Cells[outputRow, outputData_PTANo_Col].Value = ptaNo;
                                wsInput_OutputData.Cells[outputRow, outputData_DocTitle_Col].Value = docTitle;
                                wsInput_OutputData.Cells[outputRow, outputData_Purpose_Col].Value = purpose;
                                wsInput_OutputData.Cells[outputRow, outputData_Status_Col].Value = status;

                                // Tăng biến đếm và chuyển sang dòng tiếp theo trong sheet3
                                outputRow++;
                            }
                        }

                        // Định dạng Border cho toàn bộ vùng có dữ liệu
                        if (outputRow > outputData_StartRow)
                        {
                            var dataRange = wsInput_OutputData.Cells[outputData_StartRow - 1, outputData_No_Col,
                                                                   outputRow - 1, outputData_Status_Col];

                            dataRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            dataRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            dataRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            dataRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;

                            foreach (var cell in dataRange)
                            {
                                cell.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                cell.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                                cell.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                                cell.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            }
                        }

                        wsInput_OutputData.Select();
                        packageInput.Save();
                    }
                    MessageBox.Show("Done", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            #endregion
        }


    }
}