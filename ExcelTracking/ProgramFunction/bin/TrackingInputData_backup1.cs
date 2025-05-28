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
//using OfficeOpenXml;
//#endregion

//namespace SupportTools
//{
//    public class TrackingInputData
//    {
//        public static string txtFilePath_Master = "";
//        public static string txtFilePath_InputData = "";

//        public static void TransferDataFromInputToMaster()
//        {
//            if (string.IsNullOrEmpty(txtFilePath_Master) || string.IsNullOrEmpty(txtFilePath_InputData))
//            { 
//                MessageBox.Show("Vui lòng chọn đủ các file excel!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
//                return; 
//            }

//            // ========================== Process
//            // *** INPUT DATA FILE
//            // * Thông tin InputData file
//            int startRow_InputData = 11;        // Dòng bắt đầu ghi dữ liệu
//            int docRefCol = 2;                  // Cột "Doc Ref" (B)
//            int revCol_InputData = 4;           // Cột D (Rev)
//            int statusCol_InputData = 7;        // Cột G (Status)

//            // * Kiểm tra DocRef có trùng nhau ko, nếu trùng nhau thì note vào cột Status
//            // * Tạo file _Transfered cho Input Data
//            string directory = Path.GetDirectoryName(txtFilePath_InputData); // Thư mục chứa tệp gốc
//            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(txtFilePath_InputData); // Tên file không có đuôi
//            string fileExtension = Path.GetExtension(txtFilePath_InputData); // Lấy đuôi file (.xlsx)
//            string outputFilePath = Path.Combine(directory, $"{fileNameWithoutExt}_Transfered{fileExtension}");

//            FileInfo fileInfo = new FileInfo(txtFilePath_InputData);
//            if (!fileInfo.Exists)
//            {
//                MessageBox.Show("Tệp Excel không tồn tại!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
//                return;
//            }

//            // *** MASTER FILE
//            // * Thông tin Master file
//            int startRow_Master = 5;            // Dòng bắt đầu ghi dữ liệu
//            int startCol_Master_TRANSID = 38;   // Cột đầu tiên TRANS ID
//            int totalCol_Master_TRANSID = 3;    // Tổng số cột TRANS ID
//            int startCol_Master_REV = startCol_Master_TRANSID + totalCol_Master_TRANSID;
//            int startCol_Master_Purpose = startCol_Master_REV + totalCol_Master_TRANSID;
//            int allianceCol_Master = 7;         // Cột G (Alliance No.)
//            int statusCol_Master = 7;           // Cột G để ghi "Not Found"

//            // * Backup Master File
//            string masterDirectory = Path.GetDirectoryName(txtFilePath_Master);
//            string backupFolder = Path.Combine(masterDirectory, "_backup");
//            if (!Directory.Exists(backupFolder))
//            {
//                Directory.CreateDirectory(backupFolder);
//            }
//            string backupMasterPath = Path.Combine(backupFolder, Path.GetFileName(txtFilePath_Master));
//            File.Copy(txtFilePath_Master, backupMasterPath, true);




//            using (var packageMaster = new ExcelPackage(new FileInfo(txtFilePath_Master)))
//            using (var packageInput = new ExcelPackage(new FileInfo(txtFilePath_InputData)))
//            {
//                var wsMaster = packageMaster.Workbook.Worksheets[1];
//                var wsInput = packageInput.Workbook.Worksheets[1];

//                string transId = wsInput.Cells[2, 3].Text;
//                int totalRowsInput = wsInput.Dimension.End.Row;
//                int totalRowsMaster = wsMaster.Dimension.End.Row;

//                // Kiểm tra trùng lặp Doc Ref trong file InputData
//                var docRefList = wsInput.Cells[startRow_InputData, docRefCol, totalRowsInput, docRefCol]
//                                      .Select(cell => cell.Text).ToList();
//                for (int row = startRow_InputData; row <= totalRowsInput; row++)
//                {
//                    string docRef = wsInput.Cells[row, docRefCol].Text;
//                    if (!string.IsNullOrEmpty(docRef))
//                    {
//                        int count = docRefList.Count(x => x == docRef);
//                        if (count > 1)
//                        {
//                            wsInput.Cells[row, statusCol_InputData].Value = "Duplicate";
//                            wsInput.Cells[row, statusCol_InputData].Style.Font.Color.SetColor(Color.Red);
//                        }
//                        else
//                        {
//                            wsInput.Cells[row, statusCol_InputData].Value = "OK";
//                        }
//                    }
//                }

//                // Xử lý dữ liệu và đưa vào Master
//                for (int row = startRow_InputData; row <= totalRowsInput; row++)
//                {
//                    string docRef = wsInput.Cells[row, docRefCol].Text;
//                    string revValue = wsInput.Cells[row, revCol_InputData].Text;
//                    bool found = false;

//                    for (int masterRow = startRow_Master; masterRow <= totalRowsMaster; masterRow++)
//                    {
//                        string allianceNo = wsMaster.Cells[masterRow, allianceCol_Master].Text;
//                        if (allianceNo == docRef)
//                        {
//                            found = true;
//                            bool inserted = false;
//                            for (int i = 0; i < totalCol_Master_TRANSID; i++)
//                            {
//                                if (string.IsNullOrEmpty(wsMaster.Cells[masterRow, startCol_Master_TRANSID + i].Text))
//                                {
//                                    wsMaster.Cells[masterRow, startCol_Master_TRANSID + i].Value = transId;
//                                    wsMaster.Cells[masterRow, startCol_Master_REV + i].Value = revValue;
//                                    inserted = true;
//                                    break;
//                                }
//                            }

//                            if (!inserted)
//                            {
//                                wsInput.Cells[2, 4].Value = "THIẾU CỘT ĐIỀN THÔNG TIN";
//                            }
//                            break;
//                        }
//                    }

//                    if (!found)
//                    {
//                        wsInput.Cells[row, statusCol_InputData].Value = "Not Found";
//                        wsInput.Cells[row, statusCol_InputData].Style.Font.Color.SetColor(Color.Red);
//                    }
//                }

//                packageMaster.Save();
//                packageInput.SaveAs(new FileInfo(outputFilePath));
//            }










//            using (var package = new ExcelPackage(fileInfo))
//            {
//                var worksheet = package.Workbook.Worksheets[1]; // Chọn sheet đầu tiên

//                // Lấy số dòng cuối cùng có dữ liệu
//                int totalRows = worksheet.Dimension.End.Row;

//                // Lưu danh sách giá trị Doc Ref để kiểm tra trùng lặp
//                var docRefList = worksheet.Cells[startRow_InputData, docRefCol, totalRows, docRefCol]
//                                  .Select(cell => cell.Text).ToList();

//                for (int row = startRow_InputData; row <= totalRows; row++)
//                {
//                    string docRef = worksheet.Cells[row, docRefCol].Text;

//                    if (!string.IsNullOrEmpty(docRef))
//                    {
//                        int count = docRefList.Count(x => x == docRef);

//                        if (count > 1)
//                        {
//                            worksheet.Cells[row, statusCol_InputData].Value = "Duplicate";
//                            worksheet.Cells[row, statusCol_InputData].Style.Font.Color.SetColor(Color.Red); // Tô màu đỏ chữ
//                        }  
//                        else
//                        {
//                            worksheet.Cells[row, statusCol_InputData].Value = "OK";
//                        }   
//                    }
//                }

//                // Lưu file đã xử lý
//                package.SaveAs(new FileInfo(outputFilePath));
//            }




//            // *** Đưa dữ liệu từ InputData vào Master file




//        }




//    }

//    public class Support
//    {
//        public static void CopyFiles()
//        {
//            try
//            {
//                string sourcePath = @"\\srvprd2\Current Projects\LOR-Laing O'Rourke\LOR98 Byford Rail\0-Input\Structure\20240606_Precast Planks_Revizto\ST170 _ Planks _ Permacast";
//                string destinationPath_R005 = @"C:\Users\VuongLDT\Desktop\Export\R005";
//                string destinationPath_R009 = @"C:\Users\VuongLDT\Desktop\Export\R009";

//                string[] directories = Directory.GetDirectories(sourcePath, "*", SearchOption.AllDirectories);

//                foreach (string directory in directories)
//                {
//                    // Tìm và sao chép các file có ký tự _005
//                    string[] files_R005 = Directory.GetFiles(directory, "*_R005*");
//                    foreach (string file in files_R005)
//                    {
//                        string fileName = Path.GetFileName(file);
//                        string newFileName = GetNewFileName(fileName);
//                        string destFile = Path.Combine(destinationPath_R005, newFileName);
//                        File.Copy(file, destFile, true); // true để ghi đè nếu file đã tồn tại
//                    }

//                    string[] files_R009 = Directory.GetFiles(directory, "*_R009*");
//                    foreach (string file in files_R009)
//                    {
//                        string fileName = Path.GetFileName(file);
//                        string newFileName = GetNewFileName(fileName);
//                        string destFile = Path.Combine(destinationPath_R009, newFileName);
//                        File.Copy(file, destFile, true); // true để ghi đè nếu file đã tồn tại
//                    }
//                }
//                MessageBox.Show("Xong");

//            }
//            catch (Exception ex)
//            {
//                MessageBox.Show(ex.Message.ToString() + "\n==> Please Close all Revit and try again!", "Error");
//            }
//        }

//        static string GetNewFileName(string fileName)
//        {
//            // Tách phần tên và phần mở rộng của file
//            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(fileName);
//            string extension = Path.GetExtension(fileName);

//            // Tách tên file bằng dấu gạch ngang và khoảng trắng
//            string[] parts = fileNameWithoutExtension.Split(new char[] { '-', ' ', '_' }, StringSplitOptions.RemoveEmptyEntries);

//            // Kiểm tra và lấy ký tự thứ 5 và 6 từ phần tên đầu tiên
//            if (parts.Length > 0 && parts.Length >= 8)
//            {
//                bool isST = false;
//                foreach (string str in parts)
//                {
//                    if (str == "ST")
//                    {
//                        isST = true;
//                        break;
//                    }
//                }

//                string chars7 = parts[6]; // Lấy ký tự thứ 5 và 6 (0-based index)
//                string chars8 = parts[7]; // Lấy ký tự thứ 5 và 6 (0-based index)

//                if (!isST)
//                {
//                    chars7 = parts[5];
//                    chars8 = parts[6];
//                }
//                string chars78 = chars7 + "-" + chars8;

//                // Tạo tên file mới
//                string newFileNameWithoutExtension = chars78 + "_" + fileNameWithoutExtension;

//                // Trả về tên file mới với phần mở rộng
//                return newFileNameWithoutExtension + extension;
//            }

//            // Nếu không thể lấy ký tự thứ 5 và 6, trả về tên file gốc
//            return fileName;
//        }

//    }
//}