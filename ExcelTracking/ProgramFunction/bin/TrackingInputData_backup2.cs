//#region using
//using System;
//using System.Collections.Generic;
//using System.ComponentModel;
//using System.Drawing;
//using System.Globalization;
//using System.IO;
//using System.Linq;
//using System.Windows;
//using System.Windows.Forms;
//using SupportTools;
//using ExcelDataManager;
//using OfficeOpenXml;
//using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
//#endregion

//namespace ExcelDataManager
//{
//    // *** Khai báo thông tin dòng cột các file excel

//    public static class InputExcelData
//    {
//        //--------------------------------------------
//        // *** INPUT DATA FILE
//        // * Thông tin InputData file
//        public const int TRANSID_Source_Row = 2;     // Hàng 2, chứa giá trị TRANS ID
//        public const int TRANSID_Source_Col = 3;     // Cột 3, chứa giá trị TRANS ID
//        public const int DateReceive_Source_Row = 3;        // Hàng 3, chứa giá trị Date Receive
//        public const int DateReceive_Source_Col = 3;        // Cột 3, chứa giá trị Date Receive

//        public const int Start_Row = 11;            // Dòng bắt đầu đọc/ghi dữ liệu trong Input file
//        public const int DocRef_Col = 2;            // Cột "Doc Ref" (B)
//        public const int DocTitle_Col = 3;          // Cột Doc Title (C) trong Input file
//        public const int REV_Col = 4;               // Cột Rev (D)
//        public const int Purpose_Col = 11;          // Cột Purpose (K)
//        public const int TRANSID_Col = 12;          // Cột TRANS ID (L) trong Input file
//        public const int DateReceive_Col = 13;      // Cột Date Receive (M) trong Input file
//        public const int TimesCount_Col = 14;            // Cột Total Times Count trong Input file
//        public const int Status_Col = 14;           // Cột Status (N) trong Input file
//    }
//    public static class InputRecordMasterExcelData
//    {
//        //--------------------------------------------
//        // *** INPUT RECORD MASTER FILE
//        public const int Start_Row = 11;        // Dòng bắt đầu đọc/ghi dữ liệu trong Input file
//    }
//    public static class MasterExcelData
//    {
//        // *** MASTER FILE
//        // * Unique code vị trí các cột điền value
//        public const string UniqueCode_TRANSID_StartCol = "mdl01#transid";
//        public const string UniqueCode_REV_StartCol = "mdl01#stampid";
//        public const string UniqueCode_Purpose_StartCol = "mdl01#purpose";
//        public const string UniqueCode_DateReceive_StartCol = "mdl01#date";
//        public const string UniqueCode_EndCol = "mdl01#end";

//        // * Thông tin các cột Master file
//        public const string SheetName_Drawing = "DRAWINGS";     // Tên Sheet Drawing sẽ ghi dữ liệu trong Master file
//        public const string SheetName_Model = "xxx";            // Tên Sheet Model sẽ ghi dữ liệu trong Master file
//        public const int Start_Row = 6;        // Dòng bắt đầu ghi dữ liệu trong Master file
//        public const int Alliance_Col = 7;     // Cột Alliance No. (G) trong Master file
//        public const int DocTitle_Col = 8;     // Cột Doc Title (H) trong Master file
//    }
//}

//namespace SupportTools
//{
//    public class TrackingInputData
//    {
//        public static string txtFilePath_Master = "";
//        public static string txtFilePath_InputData = "";
//        public static string txtFilePath_InputDataMaster = "";
//        public static bool IsWriteToMasterFile = false;
//        public static string status_DuplicateDocRef = "Duplicate DocRef";

//        public static void CheckInputDataFile()
//        {
//            if (string.IsNullOrEmpty(txtFilePath_Master) || string.IsNullOrEmpty(txtFilePath_InputData) || string.IsNullOrEmpty(txtFilePath_InputDataMaster))
//            { 
//                MessageBox.Show("Vui lòng chọn đủ các file excel!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
//                return; 
//            }

//            // ========================== Process

//            //--------------------------------------------
//            // *** MASTER FILE
//            // * Unique code vị trí các cột điền value
//            //string startCol_UniqueCode_Master_TRANSID = "mdl01#transid";
//            //string startCol_UniqueCode_Master_REV = "mdl01#stampid";
//            //string startCol_UniqueCode_Master_Purpose = "mdl01#purpose";
//            //string startCol_UniqueCode_Master_Date = "mdl01#date";
//            //string endCol_UniqueCode = "mdl01#end";

//            // * Thông tin các cột Master file
//            //string sheetName_Master = "DRAWINGS";  // Tên Sheet sẽ ghi dữ liệu trong Master file
//            //int startRow_Master = 6;        // Dòng bắt đầu ghi dữ liệu trong Master file
//            //int allianceCol_Master = 7;     // Cột Alliance No. (G) trong Master file
//            //int docTitleCol_Master = 8;     // Cột Doc Title (H) trong Master file

//            //--------------------------------------------
//            // *** INPUT RECORD MASTER FILE
//            //int startRow_InputRecordMaster = 11;        // Dòng bắt đầu đọc/ghi dữ liệu trong Input file

//            //--------------------------------------------
//            // *** INPUT DATA FILE
//            // * Thông tin InputData file
//            //int transIdValue_Row_InputData = 2;     // Hàng 2, chứa giá trị TRANS ID
//            //int transIdValue_Col_InputData = 3;     // Cột 3, chứa giá trị TRANS ID
//            //int dateValue_Row_InputData = 3;        // Hàng 3, chứa giá trị Date Receive
//            //int dateValue_Col_InputData = 3;        // Cột 3, chứa giá trị Date Receive

//            //int startRow_InputData = 11;            // Dòng bắt đầu đọc/ghi dữ liệu trong Input file
//            //int docRefCol_InputData = 2;            // Cột "Doc Ref" (B)
//            //int docTitleCol_InputData = 3;          // Cột Doc Title (C) trong Input file
//            //int revCol_InputData = 4;               // Cột Rev (D)
//            //int purposeCol_InputData = 11;          // Cột Purpose (K)
//            //int transIdCol_InputData = 12;          // Cột TRANS ID (L) trong Input file
//            //int dateReceiveCol_InputData = 13;      // Cột Date Receive (M) trong Input file
//            //int timesCol_InputData = 14;            // Cột Total Times Count trong Input file
//            //int statusCol_InputData = 14;           // Cột Status (N) trong Input file

//            // * Tạo file _Transfered cho Input Data
//            string directory = Path.GetDirectoryName(txtFilePath_InputData); // Thư mục chứa tệp gốc
//            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(txtFilePath_InputData); // Tên file không có đuôi
//            string fileExtension = Path.GetExtension(txtFilePath_InputData); // Lấy đuôi file (.xlsx)
//            string outputFilePath = Path.Combine(directory, $"{fileNameWithoutExt}_Transfered{fileExtension}");
//            if (!IsWriteToMasterFile) outputFilePath = Path.Combine(directory, $"{fileNameWithoutExt}_Checked{fileExtension}");

//            FileInfo fileInfo = new FileInfo(txtFilePath_InputData);
//            if (!fileInfo.Exists)
//            {
//                MessageBox.Show("Tệp Excel không tồn tại!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
//                return;
//            }

//            //--------------------------------------------
//            // *** Backup Master và InputMaster trước khi transfer data
//            // * Backup Master File
//            string masterDirectory = Path.GetDirectoryName(txtFilePath_Master);
//            string backupFolder = Path.Combine(masterDirectory, "_backup_Master");
//            if (!Directory.Exists(backupFolder))
//            {
//                Directory.CreateDirectory(backupFolder);
//            }
//            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
//            string backupMasterPath = Path.Combine(backupFolder, $"{Path.GetFileNameWithoutExtension(txtFilePath_Master)}_{timestamp}{Path.GetExtension(txtFilePath_Master)}");
//            if (IsWriteToMasterFile)
//            {
//                File.Copy(txtFilePath_Master, backupMasterPath, true);
//            }
//            // * Backup Input Master File
//            string inputMasterDirectory = Path.GetDirectoryName(txtFilePath_InputDataMaster);
//            string backupInputMasterFolder = Path.Combine(inputMasterDirectory, "_backup_InputMaster");
//            if (!Directory.Exists(backupInputMasterFolder))
//            {
//                Directory.CreateDirectory(backupInputMasterFolder);
//            }
//            string backupInputMasterPath = Path.Combine(backupInputMasterFolder, $"{Path.GetFileNameWithoutExtension(txtFilePath_InputDataMaster)}_{timestamp}{Path.GetExtension(txtFilePath_InputDataMaster)}");
//            if (IsWriteToMasterFile)
//            {
//                File.Copy(txtFilePath_InputDataMaster, backupInputMasterPath, true);
//            }

//            //--------------------------------------------
//            using (var packageMaster = new ExcelPackage(new FileInfo(txtFilePath_Master)))
//            using (var packageInput = new ExcelPackage(new FileInfo(txtFilePath_InputData)))
//            using (var packageInputMaster = new ExcelPackage(new FileInfo(txtFilePath_InputDataMaster)))
//            {
//                // * Ktra có Sheet cần tìm trong Master file ko rồi mới bắt đầu
//                var wsMaster = GetWorksheetByName(packageMaster, MasterExcelData.SheetName_Drawing);
//                if (wsMaster == null)
//                {
//                    MessageBox.Show("Không tìm thấy sheet " + MasterExcelData.SheetName_Drawing + " trong Master file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
//                    return;
//                }
//                // * Kiểm tra số cột tương ứng với TRANS ID có bằng nhau ko thì mới bắt đầu
//                bool isSpacingEqual = CheckEqualColumnSpacingInRange(wsMaster, MasterExcelData.UniqueCode_TRANSID_StartCol, MasterExcelData.UniqueCode_EndCol);
//                if (!isSpacingEqual)
//                {
//                    MessageBox.Show("Số cột data không đồng đều!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
//                    return;
//                }

//                var wsInput = packageInput.Workbook.Worksheets[1];
//                var wsInputMaster = packageInputMaster.Workbook.Worksheets[1];

//                // * Thông tin Master file
//                int startCol_Master_TRANSID = GetColumnIndexByUniqueCode(wsMaster, MasterExcelData.UniqueCode_TRANSID_StartCol);     // Cột đầu tiên TRANS ID
//                int totalCol_Master_TRANSID = GetColumnCountBetweenUniqueCodes(wsMaster, MasterExcelData.UniqueCode_TRANSID_StartCol);    // Tổng số cột TRANS ID
//                int startCol_Master_REV = GetColumnIndexByUniqueCode(wsMaster, MasterExcelData.UniqueCode_REV_StartCol);
//                int startCol_Master_Purpose = GetColumnIndexByUniqueCode(wsMaster, MasterExcelData.UniqueCode_Purpose_StartCol);
//                int startCol_Master_Date = GetColumnIndexByUniqueCode(wsMaster, MasterExcelData.UniqueCode_DateReceive_StartCol);

//                // * InputData file
//                string transId = wsInput.Cells[InputExcelData.TRANSID_Source_Row, InputExcelData.TRANSID_Source_Col].Text;        // Lấy giá trị TRANS ID - Input file
//                string dateValue_str = wsInput.Cells[InputExcelData.DateReceive_Source_Row, InputExcelData.DateReceive_Source_Col].Text;        // Lấy giá trị Date - Input file

//                int totalRowsInput = wsInput.Dimension.End.Row;
//                int totalRowsMaster = wsMaster.Dimension.End.Row;
//                int totalRowsInputMaster = wsInputMaster.Dimension.End.Row;

//                // Phần 1: Kiểm tra trùng lặp Doc Ref trong file InputData
//                var docRefList = wsInput.Cells[InputExcelData.Start_Row, InputExcelData.DocRef_Col, totalRowsInput, InputExcelData.DocRef_Col]
//                                      .Select(cell => cell.Text).ToList();
//                for (int row = InputExcelData.Start_Row; row <= totalRowsInput; row++)
//                {
//                    string docRef = wsInput.Cells[row, InputExcelData.DocRef_Col].Text;
//                    if (!string.IsNullOrEmpty(docRef))
//                    {
//                        int count = docRefList.Count(x => x == docRef);
//                        if (count > 1)
//                        {
//                            wsInput.Cells[row, InputExcelData.Status_Col].Value = status_DuplicateDocRef;
//                            wsInput.Cells[row, InputExcelData.Status_Col].Style.Font.Color.SetColor(Color.Red);
//                        }
//                        else
//                        {
//                            wsInput.Cells[row, InputExcelData.Status_Col].Value = "OK";
//                        }

//                        // Gán TRANS ID vào cột TRANS ID cho tất cả DocRef trong InputData
//                        wsInput.Cells[row, InputExcelData.TRANSID_Col].Value = transId;
//                        // Gán Date Receive vào cột Date Receive cho tất cả DocRef trong InputData
//                        wsInput.Cells[row, InputExcelData.DateReceive_Col].Value = dateValue_str;
//                    }
//                }

//                // Phần 2: Kiểm tra Doc Title trong Input file và trong Master file,
//                // nếu giống thì OK, nếu khác thì note vào cột DocTitle trong Input file là "Updated to Master"
//                for (int row = InputExcelData.Start_Row; row <= totalRowsInput; row++)
//                {
//                    string docRef = wsInput.Cells[row, InputExcelData.DocRef_Col].Text;      // Lấy giá trị Doc Ref
//                    string docTitleInput = wsInput.Cells[row, InputExcelData.DocTitle_Col].Text;  // Lấy giá trị Doc Title từ Input file

//                    if (!string.IsNullOrEmpty(docRef) && !string.IsNullOrEmpty(docTitleInput))
//                    {
//                        for (int masterRow = MasterExcelData.Start_Row; masterRow <= totalRowsMaster; masterRow++)
//                        {
//                            string allianceNo = wsMaster.Cells[masterRow, MasterExcelData.Alliance_Col].Text;  // Lấy giá trị Alliance No trong Master
//                            string docTitleMaster = wsMaster.Cells[masterRow, MasterExcelData.DocTitle_Col].Text;  // Lấy giá trị Doc Title từ Master file

//                            if (allianceNo == docRef) // Nếu Doc Ref trùng với Alliance No
//                            {
//                                if (!docTitleInput.Equals(docTitleMaster, StringComparison.OrdinalIgnoreCase))
//                                {
//                                    wsInput.Cells[row, InputExcelData.DocTitle_Col].Value = "Updated to Master";
//                                    wsInput.Cells[row, InputExcelData.DocTitle_Col].Style.Font.Color.SetColor(Color.Blue);
//                                }
//                                else
//                                {
//                                    wsInput.Cells[row, InputExcelData.DocTitle_Col].Value = "OK";
//                                }
//                                break;
//                            }
//                        }
//                    }
//                }

//                // Phần 3: Xử lý dữ liệu và đưa vào Master
//                for (int row = InputExcelData.Start_Row; row <= totalRowsInput; row++)
//                {
//                    if (wsInput.Cells[row, InputExcelData.Status_Col].Text == status_DuplicateDocRef)
//                    {
//                        continue; // Bỏ qua dòng Duplicate
//                    }

//                    string docRef = wsInput.Cells[row, InputExcelData.DocRef_Col].Text;                             // Lấy giá trị Doc Ref trong Input file
//                    string revValue = wsInput.Cells[row, InputExcelData.REV_Col].Text;                    // Lấy giá trị REV trong Input file
//                    string purposeValue = wsInput.Cells[row, InputExcelData.Purpose_Col].Text;            // Lấy giá trị Purpose trong Input file
//                    string dateReceiveValue = wsInput.Cells[row, InputExcelData.DateReceive_Col].Text;    // Lấy giá trị Purpose trong Input file
//                    string transIdValue = wsInput.Cells[row, InputExcelData.TRANSID_Col].Text;            // Lấy TRANS ID từ cột K Input file

//                    bool found = false;

//                    for (int masterRow = MasterExcelData.Start_Row; masterRow <= totalRowsMaster; masterRow++)
//                    {
//                        string allianceNo = wsMaster.Cells[masterRow, MasterExcelData.Alliance_Col].Text;
//                        if (allianceNo == docRef)
//                        {
//                            found = true;
//                            bool inserted = false;

//                            // Kiểm tra nếu TRANS ID đã tồn tại trong Master
//                            bool transIdExists = false;
//                            for (int i = 0; i < totalCol_Master_TRANSID; i++)
//                            {
//                                if (wsMaster.Cells[masterRow, startCol_Master_TRANSID + i].Text == transId)
//                                {
//                                    transIdExists = true;
//                                    break;
//                                }
//                            }
//                            if (transIdExists)
//                            {
//                                wsInput.Cells[row, InputExcelData.Status_Col].Value = "Duplicate TRANS ID";
//                                wsInput.Cells[row, InputExcelData.Status_Col].Style.Font.Color.SetColor(Color.Red);
//                                break; // Bỏ qua nếu TRANS ID đã tồn tại
//                            }

//                            for (int i = 0; i < totalCol_Master_TRANSID; i++)
//                            {
//                                if (string.IsNullOrEmpty(wsMaster.Cells[masterRow, startCol_Master_TRANSID + i].Text))
//                                {
                                    
//                                    // TRANS ID
//                                    wsMaster.Cells[masterRow, startCol_Master_TRANSID + i].Value = transIdValue;
//                                    // REV
//                                    wsMaster.Cells[masterRow, startCol_Master_REV + i].Value = revValue;
//                                    // PURPOSE
//                                    wsMaster.Cells[masterRow, startCol_Master_Purpose + i].Value = purposeValue;
//                                    // DATE
//                                    DateTime dateValue;
//                                    if (DateTime.TryParseExact(dateReceiveValue, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out dateValue))
//                                    {
//                                        wsMaster.Cells[masterRow, startCol_Master_Date + i].Value = dateValue;
//                                        wsMaster.Cells[masterRow, startCol_Master_Date + i].Style.Numberformat.Format = "dd/MM/yyyy"; // Định dạng ô theo ngày tháng năm
//                                    }
//                                    else
//                                    {
//                                        wsMaster.Cells[masterRow, startCol_Master_Date + i].Value = dateReceiveValue; // Giữ nguyên nếu không parse được
//                                    }

                                    
//                                    inserted = true;
//                                    break;
//                                }
//                            }

//                            if (!inserted)
//                            {
//                                wsInput.Cells[row, InputExcelData.Status_Col].Value = "Missing column in the Master file";
//                                wsInput.Cells[row, InputExcelData.Status_Col].Style.Font.Color.SetColor(Color.Red);
//                            }
//                            break;
//                        }
//                    }

//                    if (!found)
//                    {
//                        wsInput.Cells[row, InputExcelData.Status_Col].Value = "Not Found DocRef in the Matster file";
//                        wsInput.Cells[row, InputExcelData.Status_Col].Style.Font.Color.SetColor(Color.Red);
//                    }
//                }

//                // Phần 4: Đưa data từ InputData vào Input Master

//                // Dictionary lưu số lần xuất hiện của DocRef trong Input Master
//                Dictionary<string, int> docRefCounts = new Dictionary<string, int>();

//                // Bước 1: Quét toàn bộ Input Master trước, lưu số lần xuất hiện vào Dictionary
//                for (int rowMaster = InputRecordMasterExcelData.Start_Row; rowMaster <= totalRowsInputMaster; rowMaster++)
//                {
//                    string docRefMaster = wsInputMaster.Cells[rowMaster, InputExcelData.DocRef_Col].Text;

//                    if (!string.IsNullOrEmpty(docRefMaster))
//                    {
//                        if (docRefCounts.ContainsKey(docRefMaster))
//                            docRefCounts[docRefMaster]++;
//                        else
//                            docRefCounts[docRefMaster] = 1;
//                    }
//                }

//                // Bước 2: Duyệt từng dòng trong Input Data, cập nhật số lần xuất hiện
//                for (int rowInput = InputExcelData.Start_Row; rowInput <= totalRowsInput; rowInput++)
//                {
//                    string docRef = wsInput.Cells[rowInput, InputExcelData.DocRef_Col].Text;

//                    if (!string.IsNullOrEmpty(docRef))
//                    {
//                        // Nếu đã có trong Dictionary, tăng count lên 1 (lần tiếp theo)
//                        if (docRefCounts.ContainsKey(docRef))
//                        {
//                            docRefCounts[docRef]++;
//                        }
//                        else
//                        {
//                            docRefCounts[docRef] = 1; // Lần đầu tiên xuất hiện
//                        }

//                        // Ghi giá trị Times Count vào cột Times trong Input Data
//                        wsInput.Cells[rowInput, InputExcelData.TimesCount_Col].Value = docRefCounts[docRef];
//                    }
//                }


//                // Phần cuối: Lưu các file
//                if (IsWriteToMasterFile)
//                {
//                    packageMaster.Save();
//                    packageInputMaster.Save();
//                }
//                packageInput.SaveAs(new FileInfo(outputFilePath));


//                MessageBox.Show("Done", "Update Master", MessageBoxButtons.OK, MessageBoxIcon.Information);
//            }
//        }

//        public static void TransferDataFromInputToMaster()
//        {

//        }

//        // ============================================================================================================
//        public static ExcelWorksheet GetWorksheetByName(ExcelPackage package, string sheetName)
//        {
//            return package.Workbook.Worksheets.FirstOrDefault(ws => ws.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase));
//        }

//        public static int GetColumnIndexByUniqueCode(ExcelWorksheet ws, string uniqueCode)
//        {
//            int totalColumns = ws.Dimension.End.Column; // Tổng số cột trong sheet
//            int headerRow = 1; // Dòng chứa mã unique

//            for (int col = 1; col <= totalColumns; col++)
//            {
//                if (ws.Cells[headerRow, col].Text.Trim() == uniqueCode)
//                {
//                    return col; // Trả về chỉ số cột nếu tìm thấy
//                }
//            }
//            return -1; // Trả về -1 nếu không tìm thấy
//        }
//        public static int GetColumnCountBetweenUniqueCodes(ExcelWorksheet ws, string startUniqueCode)
//        {
//            int totalColumns = ws.Dimension.End.Column; // Tổng số cột trong sheet
//            int headerRow = 1; // Dòng chứa mã unique
//            int startCol = GetColumnIndexByUniqueCode(ws, startUniqueCode);

//            if (startCol == -1) return 0; // Không tìm thấy unique code

//            int count = 1; // Bắt đầu đếm từ cột chứa unique code
//            for (int col = startCol + 1; col <= totalColumns; col++)
//            {
//                string cellValue = ws.Cells[headerRow, col].Text.Trim();
//                if (!string.IsNullOrEmpty(cellValue)) break; // Gặp unique code tiếp theo thì dừng
//                count++;
//            }
//            return count; // Số cột TRANS ID
//        }

//        public static bool CheckEqualColumnSpacingInRange(ExcelWorksheet ws, string startUniqueCode, string endUniqueCode)
//        {
//            int totalColumns = ws.Dimension.End.Column; // Tổng số cột trong sheet
//            int headerRow = 1; // Dòng chứa unique code

//            // Tìm vị trí cột của unique code đầu và cuối
//            int startCol = GetColumnIndexByUniqueCode(ws, startUniqueCode);
//            int endCol = GetColumnIndexByUniqueCode(ws, endUniqueCode); 

//            if (startCol == -1 || endCol == 0 || startCol >= endCol || endCol > totalColumns)
//            {
//                return false; // Không tìm thấy unique code hoặc vị trí không hợp lệ
//            }

//            List<int> uniqueCodePositions = new List<int>();

//            // Lưu vị trí các unique code trong khoảng chỉ định
//            for (int col = startCol; col <= endCol; col++)
//            {
//                if (!string.IsNullOrEmpty(ws.Cells[headerRow, col].Text.Trim()))
//                {
//                    uniqueCodePositions.Add(col);
//                }
//            }

//            if (uniqueCodePositions.Count < 2) return true; // Nếu chỉ có 1 unique code, mặc định là đúng

//            // Kiểm tra khoảng cách giữa các unique code
//            int expectedSpacing = uniqueCodePositions[1] - uniqueCodePositions[0];
//            for (int i = 1; i < uniqueCodePositions.Count - 1; i++)
//            {
//                int currentSpacing = uniqueCodePositions[i + 1] - uniqueCodePositions[i];
//                // Nếu là spacing cuối cùng, thì +1
//                if (i == uniqueCodePositions.Count - 2)
//                {
//                    currentSpacing += 1;
//                }
//                if (currentSpacing != expectedSpacing)
//                {
//                    return false;
//                }
//            }

//            return true;
//        }




//    }
//}