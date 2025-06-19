using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml;
using System.IO;
using System.ComponentModel;
using ExcelDataManager;
using ExcelTracking;

public class RFIMasterProcessor
{
    #region Column Definitions - Khai báo cột để dễ thay đổi

    // === RFI FILE COLUMNS ===
    private const int RFI_REF_NO_COL = 1;           // Cột A - Ref No
    private const int RFI_DOC_REF_COL = 4;          // Cột D - Doc Ref
    private const int RFI_REFERENCE_DOC_COL = 6;    // Cột F - Reference Document
    private const int RFI_BRE_ANSWER_COL = 11;      // Cột K - BRE Answer
    private const int RFI_STATUS_COL = 14;          // Cột N - Status

    // === MASTER FILE COLUMNS ===
    private const int MASTER_ALLIANCE_NO_COL = 7;   // Cột G - Alliance No
    private const int MASTER_REF_NO_LIST_COL = 69;  // Cột BQ - Danh sách Ref No
    private const int MASTER_BRE_COUNT_COL = 70;    // Cột BR - Count BRE Answer
    private const int MASTER_OPEN_COUNT_COL = 71;   // Cột BS - Count Open
    private const int MASTER_CLOSED_COUNT_COL = 72; // Cột BT - Count Closed
    private const int MASTER_OVERALL_STATUS_COL = 73; // Cột BU - Overall Status
    private const int MASTER_OPEN_ITEMS_COL = 74;   // Cột BV - Open Items Text

    #endregion

    public void ProcessRFIAndMasterFiles(string rfiFilePath, string masterFilePath)
    {
        using (var rfiPackage = new ExcelPackage(new FileInfo(rfiFilePath)))
        using (var masterPackage = new ExcelPackage(new FileInfo(masterFilePath)))
        {
            var rfiWorksheet = rfiPackage.Workbook.Worksheets[1];
            var masterWorksheet = masterPackage.Workbook.Worksheets[2];
            masterWorksheet = MainFunction.GetWorksheetByName(masterPackage, MasterExcelData_Drawing.SheetName);

            // ✅ Tập hợp để lưu các DocRef đã xử lý (tránh trùng lặp)
            var processedDocRefs = new HashSet<string>();

            // 🔄 Duyệt qua từng dòng trong RFI file
            int rfiRowCount = rfiWorksheet.Dimension.End.Row;

            for (int rfiRow = 2; rfiRow <= rfiRowCount; rfiRow++) // Bỏ qua header
            {
                var docRef = rfiWorksheet.Cells[rfiRow, RFI_DOC_REF_COL].Text?.Trim();

                // 🚫 Bỏ qua nếu DocRef rỗng hoặc đã xử lý
                if (string.IsNullOrEmpty(docRef) || processedDocRefs.Contains(docRef))
                    continue;

                // ✅ Đánh dấu DocRef đã xử lý
                processedDocRefs.Add(docRef);

                // 🔍 Tìm dòng tương ứng trong Master file
                int masterRow = FindMasterRowByAllianceNo(masterWorksheet, docRef);

                if (masterRow == -1)
                    continue; // Không tìm thấy Alliance No tương ứng

                // 📊 Xử lý tất cả dòng RFI có cùng DocRef
                ProcessDocRefGroup(rfiWorksheet, masterWorksheet, docRef, masterRow);
            }

            // 💾 Lưu file Master
            masterPackage.Save();
        }
    }

    /// <summary>
    /// 🔍 Tìm dòng trong Master file có Alliance No tương ứng với DocRef
    /// </summary>
    private int FindMasterRowByAllianceNo(ExcelWorksheet masterWorksheet, string docRef)
    {
        int masterRowCount = masterWorksheet.Dimension.End.Row;

        for (int row = 2; row <= masterRowCount; row++)
        {
            var allianceNo = masterWorksheet.Cells[row, MASTER_ALLIANCE_NO_COL].Text?.Trim();
            if (string.Equals(allianceNo, docRef, StringComparison.OrdinalIgnoreCase))
            {
                return row;
            }
        }

        return -1; // Không tìm thấy
    }

    /// <summary>
    /// 📊 Xử lý nhóm các dòng RFI có cùng DocRef
    /// </summary>
    private void ProcessDocRefGroup(ExcelWorksheet rfiWorksheet, ExcelWorksheet masterWorksheet,
                                  string docRef, int masterRow)
    {
        // 📝 Tìm tất cả dòng RFI có cùng DocRef
        var matchingRows = FindAllRFIRowsWithDocRef(rfiWorksheet, docRef);

        if (!matchingRows.Any())
            return;

        // 🔢 Thu thập dữ liệu từ các dòng matching
        var refNumbers = new List<string>();
        var breAnswerCount = 0;
        var openCount = 0;
        var closedCount = 0;
        var openItemTexts = new List<string>();

        foreach (int rfiRow in matchingRows)
        {
            // ➡️ Ref No (Cột A)
            var refNo = rfiWorksheet.Cells[rfiRow, RFI_REF_NO_COL].Text?.Trim();
            if (!string.IsNullOrEmpty(refNo))
                refNumbers.Add(refNo);

            // ➡️ BRE Answer Count (Cột K)
            var breAnswer = rfiWorksheet.Cells[rfiRow, RFI_BRE_ANSWER_COL].Text?.Trim();
            if (!string.IsNullOrEmpty(breAnswer))
                breAnswerCount++;

            // ➡️ Status Count (Cột N)
            var status = rfiWorksheet.Cells[rfiRow, RFI_STATUS_COL].Text?.Trim();
            if (string.Equals(status, "Open", StringComparison.OrdinalIgnoreCase))
            {
                openCount++;

                // ➡️ Lấy text cho Open items (Cột F)
                var referenceDoc = rfiWorksheet.Cells[rfiRow, RFI_REFERENCE_DOC_COL].Text?.Trim();
                if (!string.IsNullOrEmpty(referenceDoc))
                {
                    // 🧹 Loại bỏ phần trong ngoặc
                    var cleanText = RemoveParenthesesContent(referenceDoc);
                    if (!string.IsNullOrEmpty(cleanText))
                        openItemTexts.Add(cleanText);
                }
            }
            else if (string.Equals(status, "Closed", StringComparison.OrdinalIgnoreCase))
            {
                closedCount++;
            }
        }

        // ✏️ Điền dữ liệu vào Master file
        FillMasterFileData(masterWorksheet, masterRow, refNumbers, breAnswerCount,
                          openCount, closedCount, openItemTexts);
    }

    /// <summary>
    /// 🔍 Tìm tất cả dòng RFI có cùng DocRef
    /// </summary>
    private List<int> FindAllRFIRowsWithDocRef(ExcelWorksheet rfiWorksheet, string targetDocRef)
    {
        var matchingRows = new List<int>();
        int rfiRowCount = rfiWorksheet.Dimension.End.Row;

        for (int row = 2; row <= rfiRowCount; row++)
        {
            var docRef = rfiWorksheet.Cells[row, RFI_DOC_REF_COL].Text?.Trim();
            if (string.Equals(docRef, targetDocRef, StringComparison.OrdinalIgnoreCase))
            {
                matchingRows.Add(row);
            }
        }

        return matchingRows;
    }

    /// <summary>
    /// 🧹 Loại bỏ nội dung trong ngoặc ()
    /// </summary>
    private string RemoveParenthesesContent(string input)
    {
        if (string.IsNullOrEmpty(input))
            return string.Empty;

        int openParen = input.LastIndexOf('(');
        if (openParen >= 0)
        {
            return input.Substring(0, openParen).Trim();
        }

        return input.Trim();
    }

    /// <summary>
    /// ✏️ Điền dữ liệu vào Master file
    /// </summary>
    private void FillMasterFileData(ExcelWorksheet masterWorksheet, int masterRow,
                                   List<string> refNumbers, int breAnswerCount,
                                   int openCount, int closedCount, List<string> openItemTexts)
    {
        // 📝 Cột BQ - Danh sách Ref No (phân cách bằng ;)
        if (refNumbers.Any())
        {
            masterWorksheet.Cells[masterRow, MASTER_REF_NO_LIST_COL].Value =
                string.Join(";", refNumbers.Distinct());
        }

        // 📊 Cột BR - Count BRE Answer
        masterWorksheet.Cells[masterRow, MASTER_BRE_COUNT_COL].Value = breAnswerCount;

        // 📊 Cột BS - Count Open
        masterWorksheet.Cells[masterRow, MASTER_OPEN_COUNT_COL].Value = openCount;

        // 📊 Cột BT - Count Closed  
        masterWorksheet.Cells[masterRow, MASTER_CLOSED_COUNT_COL].Value = closedCount;

        // 🎯 Cột BU - Overall Status
        string overallStatus = (openCount > 0) ? "Open" : "Closed";
        masterWorksheet.Cells[masterRow, MASTER_OVERALL_STATUS_COL].Value = overallStatus;

        // 📝 Cột BV - Open Items Text (phân cách bằng ;)
        if (openItemTexts.Any())
        {
            masterWorksheet.Cells[masterRow, MASTER_OPEN_ITEMS_COL].Value =
                string.Join(";", openItemTexts.Distinct());
        }
    }
}

// 🚀 Cách sử dụng
//public class Program
//{
//    public static void Main()
//    {
//        var processor = new RFIMasterProcessor();

//        string rfiFilePath = @"C:\path\to\your\RFI_file.xlsx";
//        string masterFilePath = @"C:\path\to\your\Master_file.xlsx";

//        try
//        {
//            processor.ProcessRFIAndMasterFiles(rfiFilePath, masterFilePath);
//            Console.WriteLine("✅ Xử lý hoàn thành!");
//        }
//        catch (Exception ex)
//        {
//            Console.WriteLine($"❌ Lỗi: {ex.Message}");
//        }
//    }
//}