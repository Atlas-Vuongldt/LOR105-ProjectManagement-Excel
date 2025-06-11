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
using MessageBox = System.Windows.Forms.MessageBox;
using Newtonsoft.Json;
#endregion

// NOTE
// Bổ sung phần đưa RedlineMarkups và PackageStamp vào Master và backup vào Master Record

namespace ExcelTracking
{
    public class MainFunction
    {
        // ============================================================================================================
        // *** Function
        public static bool IsValidExcelFiles_ForTransfer_MainTab(Form activeForm, string caption, string sheetName_InputRecorMaster)
        {
            #region
            if (string.IsNullOrEmpty(TrackingInputData.txtFilePath_Master) || 
                string.IsNullOrEmpty(TrackingInputData.txtFilePath_InputData) || 
                string.IsNullOrEmpty(TrackingInputData.txtFilePath_InputRecordMaster))
            {
                MessageBox.Show(activeForm, "Please select all the required Excel files!", "Warning",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            DialogResult dialogResult = MessageBox.Show(activeForm, "Do you want to " + caption, "Confirmation",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dialogResult == DialogResult.No) { return false; }

            if (!IsValidCheckedFileName(Path.GetFileName(TrackingInputData.txtFilePath_InputData)))
            {
                MessageBox.Show(activeForm, "Invalid Input data file, " + "\n" +
                    "\"===> Please select a InputData Checked file", "Invalid Input data file",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            if (!IsValidCheckedFileWithStatusOK(TrackingInputData.txtFilePath_InputData))
            {
                MessageBox.Show(activeForm, "Invalid Input data file, " + "\n" +
                    "===> There is data in the Input file with an Update Status that is not \"OK\"", "Invalid Input data file",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            if (!IsValidWorkSheet(TrackingInputData.txtFilePath_InputRecordMaster, sheetName_InputRecorMaster))
            {
                MessageBox.Show(activeForm, "Invalid InputRecordMaster file" + "\n" + 
                    "===> Please select the correct InputRecordMaster file", "Invalid InputRecordMaster file",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            return true;
            #endregion
        }
        public static bool IsValidExcelFiles_ForTransfer_TSTab(Form activeForm, string caption)
        {
            #region
            if (string.IsNullOrEmpty(TrackingInputData.txtFilePath_Master) || 
                string.IsNullOrEmpty(TrackingInputData.txtFilePath_TS_InputData) || 
                string.IsNullOrEmpty(TrackingInputData.txtFilePath_TS_InputRecordMaster))
            {
                MessageBox.Show(activeForm, "Please select all the required Excel files!", "Warning",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            DialogResult dialogResult = MessageBox.Show(activeForm, "Do you want to " + caption, "Confirmation",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dialogResult == DialogResult.No) { return false; }

            return true;
            #endregion
        }

        /// <summary>
        /// Lấy về cột chứa mã uniqueCode, code được hiểu mặc định ở Row = 1
        /// </summary>
        public static int GetColumnIndexByUniqueCode(string excelFilePath, string SheetName, string uniqueCode)
        {
            #region
            using (var packageExcel = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                var ws = GetWorksheetByName(packageExcel, SheetName);
                int totalColumns = ws.Dimension.End.Column; // Tổng số cột trong sheet
                int headerRow = 1; // Dòng chứa mã unique

                for (int col = 1; col <= totalColumns; col++)
                {
                    if (ws.Cells[headerRow, col].Text.Trim() == uniqueCode)
                    {
                        return col; // Trả về chỉ số cột nếu tìm thấy
                    }
                }
            }
            return -1; // Trả về -1 nếu không tìm thấy
            #endregion
        }

        public static bool IsValidCheckedFileName(string fileName)
        {
            #region
            if (!fileName.Contains("_Checked_")) return false;
            string[] parts = fileName.Split('_');
            if (parts.Length < 3) return false;

            string dateTimePart = parts[parts.Length - 2] + "_" + parts[parts.Length - 1].Replace(".xlsx", ""); // Lấy 2 phần cuối cùng chứa timestamp
            if (dateTimePart.Length < 15) return false;

            string datePart = dateTimePart.Substring(0, 8);
            string timePart = dateTimePart.Substring(9, 6);

            return DateTime.TryParseExact(datePart, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out _) &&
                   DateTime.TryParseExact(timePart, "HHmmss", CultureInfo.InvariantCulture, DateTimeStyles.None, out _);
            #endregion
        }

        public static bool IsValidCheckedFileWithStatusOK(string filePath)
        {
            #region
            int titleRow = 10;
            int startRow = 11;
            int docRef_Col = 2;

            if (!File.Exists(filePath))
                throw new FileNotFoundException("File không tồn tại.", filePath);

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = GetWorksheetByName(package, TrackingInputData.InputDataTemplate_Sheet_Drawing);

                if (worksheet == null)
                    throw new Exception("Không thể đọc dữ liệu từ file Excel.");

                int rowCount = worksheet.Dimension.Rows; // Tổng số dòng
                int colCount = worksheet.Dimension.Columns; // Tổng số cột
                int updateStatusColumnIndex = -1;

                // Xác định vị trí cột "Update Status"
                for (int col = 1; col <= colCount; col++)
                {
                    if (worksheet.Cells[titleRow, col].Text.Trim().Equals("Update Status", StringComparison.OrdinalIgnoreCase))
                    {
                        updateStatusColumnIndex = col;
                        break;
                    }
                }

                if (updateStatusColumnIndex == -1)
                    throw new Exception("Không tìm thấy cột 'Update Status' trong file Excel.");

                // Kiểm tra giá trị trong cột "Update Status"
                for (int row = startRow; row <= rowCount; row++) // Bắt đầu từ dòng 11
                {
                    string docRef_Value = worksheet.Cells[row, docRef_Col].Text.Trim();

                    if (!string.IsNullOrEmpty(docRef_Value))    // Check cột DocRef có giá trị thì mới check tiếp
                    {
                        string cellValue = worksheet.Cells[row, updateStatusColumnIndex].Text.Trim();

                        if (string.IsNullOrEmpty(cellValue) || !cellValue.Equals("OK", StringComparison.OrdinalIgnoreCase))
                        {
                            return false; // Nếu có giá trị rỗng hoặc khác "OK" thì trả về false
                        }
                    }
                }
            }
            return true; // Tất cả giá trị hợp lệ (OK)
            #endregion
        }

        public static bool IsValidWorkSheet(string fileNamePath, string sheetName)
        {
            #region
            using (var package = new ExcelPackage(new FileInfo(fileNamePath)))
            {
                var ws = GetWorksheetByName(package, sheetName);
                if (ws == null)
                {
                    return false;
                }
            }
            return true;
            #endregion
        }

        public static void TransferToInputRecordMaster(ExcelWorksheet wsInput, ExcelWorksheet wsInputRecordMaster,
            int docRef_InputRecordMaster_Col,
            int docTitle_InputRecordMaster_Col,
            int ver_InputRecordMaster_Col,
            int rev_InputRecordMaster_Col,
            int purpose_InputRecordMaster_Col,
            int status_InputRecordMaster_Col,
            int transid_InputRecordMaster_Col,
            int date_InputRecordMaster_Col,
            int timesCount_InputRecordMaster_Col,
            int package_InputRecordMaster_Col,
            int discipline_InputRecordMaster_Col,
            int redlineMarkup_InputRecordMaster_Col,
            int packageStampStatus_InputRecordMaster_Col,
            int date_Issue_InputRecordMaster_Col,
            int modelName_InputRecordMaster_Col,
            int nativeFileType_InputRecordMaster_Col,
            int submittedFileType_InputRecordMaster_Col,
            int atlasComment_InputRecordMaster_Col,
            int updateStatus_InputRecordMaster_Col,
            int checkDocTitle_InputRecordMaster_Col,
            string dateTime_Checked,
            out bool isInput_Modified)
        {
            #region
            isInput_Modified = false;

            int docRef_Col = 2;

            // * Lấy về tổng số dòng của các file
            int totalRowsInput = FindLastRowWithDataInColumn(wsInput, 2, InputExcelData.Start_Row);
            //int totalRowsInputRecordMaster = wsInputRecordMaster.Dimension.End.Row;
            int totalRowsInputRecordMaster = wsInputRecordMaster.Cells[wsInputRecordMaster.Dimension.Start.Row, docRef_Col, wsInputRecordMaster.Dimension.End.Row, docRef_Col]
                .Where(cell => !string.IsNullOrWhiteSpace(cell.Text))
                .Max(cell => cell.Start.Row);

            // *** Đưa data vào InputRecordMaster
            // Tìm dòng tiếp theo để chèn vào Input Master
            int insertRow_InputRecordMaster = totalRowsInputRecordMaster + 1;
            for (int row = InputExcelData.Start_Row; row <= totalRowsInput; row++)
            {
                string numberValue = wsInput.Cells[row, 1].Text;
                string docRefValue = wsInput.Cells[row, InputExcelData.DocRef_Col].Text;
                string docTitleValue = wsInput.Cells[row, InputExcelData.DocTitle_Col].Text;
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
                string dateIssueValue_str = wsInput.Cells[row, InputExcelData.DateIssue_Col].Text;
                string modelNameValue = wsInput.Cells[row, InputExcelData.ModelName_Col].Text;
                string nativeFileTypeValue = wsInput.Cells[row, InputExcelData.NativeFileType_Col].Text;
                string submittedFileTypeValue = wsInput.Cells[row, InputExcelData.SubmittedFileType_Col].Text;
                string atlasCommentValue = wsInput.Cells[row, InputExcelData.AtlasComment_Col].Text;
                string updateStatusValue = wsInput.Cells[row, InputExcelData.UpdateStatus_Col].Text;
                string checkDocTitleValue = wsInput.Cells[row, InputExcelData.CheckDocTitle_Col].Text;

                if (string.IsNullOrEmpty(docRefValue)) { continue; }

                if (checkDocTitleValue.StartsWith("Need to update", StringComparison.OrdinalIgnoreCase))
                {
                    checkDocTitleValue = "Updated to Master";
                    wsInput.Cells[row, InputExcelData.CheckDocTitle_Col].Value = checkDocTitleValue;
                    isInput_Modified = true;
                }

                // =============== Transfer data
                // NUMBER 
                wsInputRecordMaster.Cells[insertRow_InputRecordMaster, 1].Value = numberValue;
                if (int.TryParse(numberValue, out int numberValue_int))
                {
                    wsInputRecordMaster.Cells[insertRow_InputRecordMaster, 1].Value = numberValue_int;
                }
                // DOC REF
                wsInputRecordMaster.Cells[insertRow_InputRecordMaster, docRef_InputRecordMaster_Col].Value = docRefValue;
                // DOC TITLE
                wsInputRecordMaster.Cells[insertRow_InputRecordMaster, docTitle_InputRecordMaster_Col].Value = docTitleValue;
                // VER
                if (ver_InputRecordMaster_Col >= 0)
                {
                    wsInputRecordMaster.Cells[insertRow_InputRecordMaster, ver_InputRecordMaster_Col].Value = verValue;
                    if (int.TryParse(verValue, out int verValue_int))
                    {
                        wsInputRecordMaster.Cells[insertRow_InputRecordMaster, ver_InputRecordMaster_Col].Value = verValue_int;
                    }
                }
                
                // REV
                wsInputRecordMaster.Cells[insertRow_InputRecordMaster, rev_InputRecordMaster_Col].Value = revValue;
                if (int.TryParse(revValue, out int revValue_int))
                {
                    wsInputRecordMaster.Cells[insertRow_InputRecordMaster, rev_InputRecordMaster_Col].Value = revValue_int;
                }
                // PURPOSE
                wsInputRecordMaster.Cells[insertRow_InputRecordMaster, purpose_InputRecordMaster_Col].Value = purposeValue;
                // STATUS
                wsInputRecordMaster.Cells[insertRow_InputRecordMaster, status_InputRecordMaster_Col].Value = statusValue;
                // TRANS ID
                wsInputRecordMaster.Cells[insertRow_InputRecordMaster, transid_InputRecordMaster_Col].Value = transIdValue;
                // DATE RECEIVE
                DateTime dateValue_ddmmyyy;
                if (DateTime.TryParseExact(dateValue_str, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out dateValue_ddmmyyy))
                {
                    wsInputRecordMaster.Cells[insertRow_InputRecordMaster, date_InputRecordMaster_Col].Value = dateValue_ddmmyyy;
                    wsInputRecordMaster.Cells[insertRow_InputRecordMaster, date_InputRecordMaster_Col].Style.Numberformat.Format = "dd/MM/yyyy"; // Định dạng ô theo ngày tháng năm
                }
                else
                {
                    wsInputRecordMaster.Cells[insertRow_InputRecordMaster, date_InputRecordMaster_Col].Value = dateValue_str; // Giữ nguyên nếu không parse được
                }
                // DATE ISSUE
                if (date_Issue_InputRecordMaster_Col >= 0)
                {
                    if (DateTime.TryParseExact(dateIssueValue_str, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out dateValue_ddmmyyy))
                    {
                        wsInputRecordMaster.Cells[insertRow_InputRecordMaster, date_Issue_InputRecordMaster_Col].Value = dateValue_ddmmyyy;
                        wsInputRecordMaster.Cells[insertRow_InputRecordMaster, date_Issue_InputRecordMaster_Col].Style.Numberformat.Format = "dd/MM/yyyy"; // Định dạng ô theo ngày tháng năm
                    }
                    else
                    {
                        wsInputRecordMaster.Cells[insertRow_InputRecordMaster, date_Issue_InputRecordMaster_Col].Value = dateIssueValue_str; // Giữ nguyên nếu không parse được
                    }
                }
                // TIMES COUNT
                wsInputRecordMaster.Cells[insertRow_InputRecordMaster, timesCount_InputRecordMaster_Col].Value = timesCountValue;
                if (int.TryParse(timesCountValue, out int timesCountValue_int))
                {
                    wsInputRecordMaster.Cells[insertRow_InputRecordMaster, timesCount_InputRecordMaster_Col].Value = timesCountValue_int;
                }
                // PACKAGE
                if (package_InputRecordMaster_Col >= 0)
                {
                    wsInputRecordMaster.Cells[insertRow_InputRecordMaster, package_InputRecordMaster_Col].Value = packageValue;
                }
                // DISCIPLINE
                if (discipline_InputRecordMaster_Col >= 0)
                {
                    wsInputRecordMaster.Cells[insertRow_InputRecordMaster, discipline_InputRecordMaster_Col].Value = disciplineValue;
                }
                // ATLAS COMMENT
                wsInputRecordMaster.Cells[insertRow_InputRecordMaster, atlasComment_InputRecordMaster_Col].Value = atlasCommentValue;
                // UPDATE STATUS
                if (updateStatus_InputRecordMaster_Col >= 0)
                {
                    wsInputRecordMaster.Cells[insertRow_InputRecordMaster, updateStatus_InputRecordMaster_Col].Value = updateStatusValue;
                }
                // CHECK DOCTITLE
                if (checkDocTitle_InputRecordMaster_Col >= 0)
                {
                    wsInputRecordMaster.Cells[insertRow_InputRecordMaster, checkDocTitle_InputRecordMaster_Col].Value = checkDocTitleValue;
                }
                // Redline Markup (DRAWING RECEIVE)
                if (wsInputRecordMaster.Name.StartsWith("Master_RLMU_Receive_"))
                {
                    wsInputRecordMaster.Cells[insertRow_InputRecordMaster, redlineMarkup_InputRecordMaster_Col].Value = redlineMarkupValue;
                }
                // Package Stamp Status (DRAWING RECEIVE)
                if (wsInputRecordMaster.Name.StartsWith("Master_RLMU_Receive_"))
                {
                    wsInputRecordMaster.Cells[insertRow_InputRecordMaster, packageStampStatus_InputRecordMaster_Col].Value = packageStampStatusValue;
                }

                // MODEL NAME (DRAWING SUBMIT / FEEDBACK)
                if (wsInputRecordMaster.Name.StartsWith("Master_DWG_Submit_") || wsInputRecordMaster.Name.StartsWith("Master_DWG_Feedback_"))
                {
                    wsInputRecordMaster.Cells[insertRow_InputRecordMaster, modelName_InputRecordMaster_Col].Value = modelNameValue;
                }

                // Native File Type (MODEL RECEIVE / SUBMIT / FEEDBACK)
                if (wsInputRecordMaster.Name.StartsWith("Master_Model_Receive") || 
                    wsInputRecordMaster.Name.StartsWith("Master_Model_Submit") || 
                    wsInputRecordMaster.Name.StartsWith("Master_Model_Feedback"))
                {
                    wsInputRecordMaster.Cells[insertRow_InputRecordMaster, nativeFileType_InputRecordMaster_Col].Value = nativeFileTypeValue;
                }

                // Submitted File Type (MODEL RECEIVE / SUBMIT)
                if (wsInputRecordMaster.Name.StartsWith("Master_Model_Submit") || wsInputRecordMaster.Name.StartsWith("Master_Model_Feedback"))
                {
                    wsInputRecordMaster.Cells[insertRow_InputRecordMaster, submittedFileType_InputRecordMaster_Col].Value = submittedFileTypeValue;
                }

                // DATETIME CHECKED
                wsInputRecordMaster.Cells[insertRow_InputRecordMaster, 100].Value = dateTime_Checked;

                insertRow_InputRecordMaster++;
            }
            #endregion
        }

        public static void TransferToInputRecordMaster_RFI(ExcelWorksheet wsInput, ExcelWorksheet wsInputRecordMaster,
            int number_InputRecordMaster_Col,
            int docRef_InputRecordMaster_Col,
            int docTitle_InputRecordMaster_Col,
            int drawingNo_InputRecordMaster_Col,
            int dateRequest_InputRecordMaster_Col,
            int dateFeedback_InputRecordMaster_Col,
            int statusRFI_InputRecordMaster_Col,
            int timesCount_InputRecordMaster_Col)
        {
            #region
            DateTime parsedDate;
            int inputRecordMasterLastRow = wsInputRecordMaster.Dimension.End.Row + 1;
            int inputLastRow = wsInput.Dimension.End.Row;

            // Duyệt qua từng dòng của Input từ dòng 5 trở đi
            for (int row = 6; row <= inputLastRow; row++)
            {
                bool isValid = !string.IsNullOrEmpty(wsInput.Cells[row, 3].Text);

                if (isValid)
                {
                    string number = wsInput.Cells[row, InputExcelData_RFI.No_Col].Text;
                    string dateRequest_str = wsInput.Cells[row, InputExcelData_RFI.DateRequest_Col].Text;
                    string drawingNo = wsInput.Cells[row, InputExcelData_RFI.DrawingNo_Col].Text;
                    string docRef = wsInput.Cells[row, InputExcelData_RFI.DocRef_Col].Text;
                    string docTitle = wsInput.Cells[row, InputExcelData_RFI.DocTitle_Col].Text;
                    string dateFeedback_str = wsInput.Cells[row, InputExcelData_RFI.DateFeedback_Col].Text;
                    string statusRFI = wsInput.Cells[row, InputExcelData_RFI.StatusRFI_Col].Text;
                    string timesCount = wsInput.Cells[row, InputExcelData_RFI.TimesCount_Col].Text;

                    // NUMBER
                    wsInputRecordMaster.Cells[inputRecordMasterLastRow, number_InputRecordMaster_Col].Value = number;
                    if (int.TryParse(number, out int numberValue_int))
                    {
                        wsInputRecordMaster.Cells[inputRecordMasterLastRow, number_InputRecordMaster_Col].Value = numberValue_int;
                    }

                    //  DOC REF
                    wsInputRecordMaster.Cells[inputRecordMasterLastRow, docRef_InputRecordMaster_Col].Value = docRef;

                    // DOC TITLE
                    wsInputRecordMaster.Cells[inputRecordMasterLastRow, docTitle_InputRecordMaster_Col].Value = docTitle;

                    // DRAWING NUMBER
                    wsInputRecordMaster.Cells[inputRecordMasterLastRow, drawingNo_InputRecordMaster_Col].Value = drawingNo;

                    // DATE REQUEST
                    wsInputRecordMaster.Cells[inputRecordMasterLastRow, dateRequest_InputRecordMaster_Col].Value = dateRequest_str;
                    if (DateTime.TryParseExact(dateRequest_str, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out parsedDate))
                    {
                        wsInputRecordMaster.Cells[inputRecordMasterLastRow, dateRequest_InputRecordMaster_Col].Value = parsedDate;
                        wsInputRecordMaster.Cells[inputRecordMasterLastRow, dateRequest_InputRecordMaster_Col].Style.Numberformat.Format = "dd/MM/yyyy"; // Định dạng ô theo ngày tháng năm
                    }

                    // DATE FEEDBACK
                    wsInputRecordMaster.Cells[inputRecordMasterLastRow, dateFeedback_InputRecordMaster_Col].Value = dateFeedback_str;
                    if (DateTime.TryParseExact(dateFeedback_str, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out parsedDate))
                    {
                        wsInputRecordMaster.Cells[inputRecordMasterLastRow, dateFeedback_InputRecordMaster_Col].Value = parsedDate;
                        wsInputRecordMaster.Cells[inputRecordMasterLastRow, dateFeedback_InputRecordMaster_Col].Style.Numberformat.Format = "dd/MM/yyyy"; // Định dạng ô theo ngày tháng năm
                    }

                    // RFI STATUS
                    wsInputRecordMaster.Cells[inputRecordMasterLastRow, statusRFI_InputRecordMaster_Col].Value = statusRFI;

                    // TIMES COUNT
                    wsInputRecordMaster.Cells[inputRecordMasterLastRow, timesCount_InputRecordMaster_Col].Value = timesCount;
                    if (int.TryParse(timesCount, out int timesCount_int))
                    {
                        wsInputRecordMaster.Cells[inputRecordMasterLastRow, timesCount_InputRecordMaster_Col].Value = timesCount_int;
                    }

                    inputRecordMasterLastRow++;
                }
            }
            #endregion
        }
        public static void TransferToInputRecordMaster_TS(ExcelWorksheet wsInput, ExcelWorksheet wsInputRecordMaster,
            int totalHour_Input_Col,
            int wpr_InputRecordMaster_StartCol,
            string reportNo_Input_Value,
            string dateReport_Input_Value)
        {
            #region
            DateTime parsedDate;
            int inputRecordMasterLastRow = wsInputRecordMaster.Dimension.End.Row + 1;
            int inputLastRow = wsInput.Dimension.End.Row;

            // Dictionary để lưu dữ liệu InputRecordMaster
            Dictionary<string, int> recordMasterDict = new Dictionary<string, int>();
            for (int row = 6; row <= wsInputRecordMaster.Dimension.End.Row; row++)
            {
                string alliance_InputRecordMaster = wsInputRecordMaster.Cells[row, 2].Text;
                if (!string.IsNullOrEmpty(alliance_InputRecordMaster))
                {
                    recordMasterDict[alliance_InputRecordMaster] = row;
                }
            }

            // Xử lý các ô trống trong dòng 4 và 5 của InputRecordMaster từ cột J trở đi
            // Tìm cột trống tiếp theo để điền Weekly report
            int wpr_InsertValue_Col = -1;
            for (int col = wpr_InputRecordMaster_StartCol; col <= 1000; col++)
            {
                if (string.IsNullOrWhiteSpace(wsInputRecordMaster.Cells[4, col].Text))  // Tìm cột trống để điền giá trị
                {
                    wsInputRecordMaster.Cells[5, col].Value = reportNo_Input_Value;

                    // Chuyển đổi định dạng ngày tháng
                    if (DateTime.TryParseExact(dateReport_Input_Value, "ddd, MMM dd, yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out parsedDate))
                    {
                        wsInputRecordMaster.Cells[4, col].Value = parsedDate;
                        wsInputRecordMaster.Cells[4, col].Style.Numberformat.Format = "dd/MM/yyyy";
                    }
                    else
                    {
                        wsInputRecordMaster.Cells[4, col].Value = dateReport_Input_Value;
                    }
                    wpr_InsertValue_Col = col;
                    break;
                }
            }

            // Duyệt qua từng dòng của Input từ dòng 5 trở đi
            for (int row = 5; row <= inputLastRow; row++)
            {
                double totalHourValue;
                bool isValid = double.TryParse(wsInput.Cells[row, totalHour_Input_Col].Text, out totalHourValue);

                if (isValid && totalHourValue > 0)
                {
                    string package = wsInput.Cells[row, 1].Text;
                    string atlasDisc = wsInput.Cells[row, 2].Text;
                    string alliance = wsInput.Cells[row, 3].Text;

                    if (recordMasterDict.ContainsKey(alliance))
                    {
                        int existingRow = recordMasterDict[alliance];
                        wsInputRecordMaster.Cells[existingRow, wpr_InsertValue_Col].Value = totalHourValue;
                    }
                    else
                    {
                        wsInputRecordMaster.Cells[inputRecordMasterLastRow, 2].Value = alliance;
                        wsInputRecordMaster.Cells[inputRecordMasterLastRow, 3].Value = package;
                        wsInputRecordMaster.Cells[inputRecordMasterLastRow, 4].Value = atlasDisc;
                        wsInputRecordMaster.Cells[inputRecordMasterLastRow, wpr_InsertValue_Col].Value = totalHourValue;

                        // Tô đỏ ô mới thêm
                        using (var range = wsInputRecordMaster.Cells[inputRecordMasterLastRow, 2, inputRecordMasterLastRow, wpr_InsertValue_Col])
                        {
                            range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Red);
                        }
                        inputRecordMasterLastRow++;
                    }
                }
            }
            #endregion
        }

        public static ExcelWorksheet GetWorksheetByName(ExcelPackage package, string sheetName)
        {
            #region
            return package.Workbook.Worksheets.FirstOrDefault(ws => ws.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase));
            #endregion
        }
        public static void BackupFileToBackupFolder(string filePath, string folderBackupName)
        {
            #region
            string fileDirectory = Path.GetDirectoryName(filePath);
            string backupFolder = Path.Combine(fileDirectory, folderBackupName);
            if (!Directory.Exists(backupFolder))
            {
                Directory.CreateDirectory(backupFolder);
            }
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string backupFilePath = Path.Combine(backupFolder, $"{Path.GetFileNameWithoutExtension(filePath)}_{timestamp}{Path.GetExtension(filePath)}");

            File.Copy(filePath, backupFilePath, true);
            #endregion
        }

        public static int FindLastRowWithDataInColumn(ExcelWorksheet worksheet, int columnIndex, int fromRow)
        {
            #region
            // Lấy tổng số dòng theo dimension để giới hạn tìm kiếm
            int totalRows = worksheet.Dimension.End.Row;

            // Duyệt từ dòng cuối lên để tìm dòng cuối cùng có dữ liệu
            for (int row = totalRows; row >= fromRow; row--)
            {
                if (!string.IsNullOrWhiteSpace(worksheet.Cells[row, columnIndex].Text))
                {
                    return row;
                }
            }

            // Nếu không tìm thấy dòng nào có dữ liệu, trả về dòng bắt đầu
            return fromRow;
            #endregion
        }




        // ============================================================================================================
        // *** OLD Function (Not use)
        public static int GetColumnIndexByUniqueCode_Old(ExcelWorksheet ws, string uniqueCode)
        {
            #region
            int totalColumns = ws.Dimension.End.Column; // Tổng số cột trong sheet
            int headerRow = 1; // Dòng chứa mã unique

            for (int col = 1; col <= totalColumns; col++)
            {
                if (ws.Cells[headerRow, col].Text.Trim() == uniqueCode)
                {
                    return col; // Trả về chỉ số cột nếu tìm thấy
                }
            }
            return -1; // Trả về -1 nếu không tìm thấy
            #endregion
        }
        public static int GetColumnCountBetweenUniqueCodes(ExcelWorksheet ws, string startUniqueCode)
        {
            #region
            int totalColumns = ws.Dimension.End.Column; // Tổng số cột trong sheet
            int headerRow = 1; // Dòng chứa mã unique
            int startCol = GetColumnIndexByUniqueCode_Old(ws, startUniqueCode);

            if (startCol == -1) return 0; // Không tìm thấy unique code

            int count = 1; // Bắt đầu đếm từ cột chứa unique code
            for (int col = startCol + 1; col <= totalColumns; col++)
            {
                string cellValue = ws.Cells[headerRow, col].Text.Trim();
                if (!string.IsNullOrEmpty(cellValue)) break; // Gặp unique code tiếp theo thì dừng
                count++;
            }
            return count; // Số cột TRANS ID
            #endregion
        }

        public static bool CheckEqualColumnSpacingInRange(ExcelWorksheet ws, string startUniqueCode, string endUniqueCode)
        {
            #region
            int totalColumns = ws.Dimension.End.Column; // Tổng số cột trong sheet
            int headerRow = 1; // Dòng chứa unique code

            // Tìm vị trí cột của unique code đầu và cuối
            int startCol = GetColumnIndexByUniqueCode_Old(ws, startUniqueCode);
            int endCol = GetColumnIndexByUniqueCode_Old(ws, endUniqueCode);

            if (startCol == -1 || endCol == 0 || startCol >= endCol || endCol > totalColumns)
            {
                return false; // Không tìm thấy unique code hoặc vị trí không hợp lệ
            }

            List<int> uniqueCodePositions = new List<int>();

            // Lưu vị trí các unique code trong khoảng chỉ định
            for (int col = startCol; col <= endCol; col++)
            {
                if (!string.IsNullOrEmpty(ws.Cells[headerRow, col].Text.Trim()))
                {
                    uniqueCodePositions.Add(col);
                }
            }

            if (uniqueCodePositions.Count < 2) return true; // Nếu chỉ có 1 unique code, mặc định là đúng

            // Kiểm tra khoảng cách giữa các unique code
            int expectedSpacing = uniqueCodePositions[1] - uniqueCodePositions[0];
            for (int i = 1; i < uniqueCodePositions.Count - 1; i++)
            {
                int currentSpacing = uniqueCodePositions[i + 1] - uniqueCodePositions[i];
                // Nếu là spacing cuối cùng, thì +1
                if (i == uniqueCodePositions.Count - 2)
                {
                    currentSpacing += 1;
                }
                if (currentSpacing != expectedSpacing)
                {
                    return false;
                }
            }

            return true;
            #endregion
        }

        
    }
}