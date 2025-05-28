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
using static ExcelTracking.MainFunction;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
#endregion

namespace ExcelDataManager
{
    // *** Khai báo thông tin dòng cột các file excel

    //==========================================================================
    // *** INPUT DATA
    public static class InputExcelData
    {
        //--------------------------------------------
        // *** INPUT DATA FILE
        // * Thông tin InputData file
        public const int TRANSID_Source_Row = 2;     // Hàng 2, chứa giá trị TRANS ID
        public const int TRANSID_Source_Col = 3;     // Cột 3, chứa giá trị TRANS ID
        public const int DateReceive_Source_Row = 3;        // Hàng 3, chứa giá trị Date Receive
        public const int DateReceive_Source_Col = 3;        // Cột 3, chứa giá trị Date Receive

        // Row info
        public const int Start_Row = 11;            // Dòng bắt đầu đọc/ghi dữ liệu
        
        // Column Info
        public const int DocRef_Col = 2;
        public const int DocTitle_Col = 3;
        
        public const int Ver_Col = 4;
        public const int REV_Col = 5;
        public const int Purpose_Col = 6;
        public const int Status_Col = 7;
        public const int TRANSID_Col = 8;
        public const int Date_Col = 9;
        public const int TimesCount_Col = 10;
        public const int Package_Col = 11;       
        public const int Discipline_Col = 12;
        public const int NativeFileType_Col = 13;
        public const int SubmittedFileType_Col = 14;
        public const int ModelName_Col = 15;
        public const int RedlineMarkup_Col = 16;
        public const int PackageStampStatus_Col = 17;
        public const int PTANumber_Col = 18;
        public const int DateIssue_Col = 19;

        public const int AtlasComment_Col = 21;
        public const int UpdateStatus_Col = 22;
        public const int CheckDocTitle_Col = 23;
    }

    public static class InputExcelData_RFI
    {
        //--------------------------------------------
        // *** INPUT DATA FILE
        // * Thông tin InputData file
        // Row info
        public const int Start_Row = 6;            // Dòng bắt đầu đọc/ghi dữ liệu

        // Column Info
        public const int No_Col = 1;
        public const int DateRequest_Col = 2;
        public const int DrawingNo_Col = 3;
        public const int DocRef_Col = 4;
        public const int DocTitle_Col = 5;
        public const int DateFeedback_Col = 11;
        public const int StatusRFI_Col = 12;
        public const int TimesCount_Col = 13;
    }

    //==========================================================================
    // *** INPUT RECORD MASTER
    public static class InputRecordMasterExcelData
    {
        // Row info
        public const int Start_Row = 5;        // Dòng bắt đầu đọc/ghi dữ liệu
        
        // Column Info
        public const int DocRef_Col = 2;
        public const int DocTitle_Col = 3;
        public const int TRANSID_Col = 8;
    }

    // * DRAWING - WORKING
    public static class InputRecordMasterExcelData_WORKING_Receive
    {
        // * Thông tin các cột InputRecordMaster file
        public static string sheetName = "Master_WORKING_Receive";

        public const int REV_Col = 5;
        public const int Purpose_Col = 6;
        public const int Date_Receive_Col = 9;
        public const int Date_Issue_Col = 15;
        public const int TimesCount_Col = 10;
        public const int Status_Col = 7;
        public const int AtlasComment_Col = 18;
    }

    // * DRAWING - FOR FIRST
    public static class InputRecordMasterExcelData_RLMU_Receive_1st
    {
        // * Thông tin các cột InputRecordMaster file
        public static string sheetName = "Master_RLMU_Receive_1st";

        public const int Ver_Col = 4;
        public const int REV_Col = 5;
        public const int Purpose_Col = 6;
        public const int Status_Col = 7;
        public const int Date_Col = 9;
        public const int TimesCount_Col = 10;
        public const int Package_Col = 11;
        public const int Discipline_Col = 12;
        public const int RedlineMarkup_Col = 13;        // Master Drawing (Receive)
        public const int PackageStampStatus_Col = 14;   // Master Drawing (Receive)

        public const int AtlasComment_Col = 18;
        public const int UpdateStatus_Col = 19;
        public const int CheckDocTitle_Col = 20;
    }
    public static class InputRecordMasterExcelData_Drawing_Submit_1st
    {
        // * Thông tin các cột InputRecordMaster file
        public static string sheetName = "Master_DWG_Submit_1st";

        public const int Ver_Col = 4;
        public const int REV_Col = 5;
        public const int Purpose_Col = 6;
        public const int Status_Col = 7;
        public const int Date_Col = 9;
        public const int TimesCount_Col = 10;
        public const int Package_Col = 11;
        public const int Discipline_Col = 12;
        public const int ModelName_Col = 13;            // Master Drawing (Submit/Feedback)
        
        public const int AtlasComment_Col = 18;
        public const int UpdateStatus_Col = 19;
        public const int CheckDocTitle_Col = 20;
    }
    public static class InputRecordMasterExcelData_Drawing_RFI_1st
    {
        // * Thông tin các cột InputRecordMaster file
        public static string sheetName = "Master_DWG_RFI_1st";

        public const int Number_Col = 1;
        public const int DocRef_Col = 2;
        public const int DocTitle_Col = 3;
        public const int DrawingNo_Col = 4;
        public const int DateRequest_Col = 5;
        public const int DateFeedback_Col = 6;
        public const int StatusRFI_Col = 7;
        public const int TimesCount_Col = 8;
    }
    public static class InputRecordMasterExcelData_Drawing_Feedback_1st
    {
        // * Thông tin các cột InputRecordMaster file
        public static string sheetName = "Master_DWG_Feedback_1st";

        public const int Ver_Col = 4;
        public const int REV_Col = 5;
        public const int Purpose_Col = 6;
        public const int Status_Col = 7;
        public const int Date_Col = 9;
        public const int TimesCount_Col = 10;
        public const int Package_Col = 11;
        public const int Discipline_Col = 12;
        public const int ModelName_Col = 13;            // Master Drawing (Submit/Feedback)

        public const int AtlasComment_Col = 18;
        public const int UpdateStatus_Col = 19;
        public const int CheckDocTitle_Col = 20;
    }
    
    // * DRAWING - FOR FINAL
    public static class InputRecordMasterExcelData_RLMU_Receive_Final
    {
        // * Thông tin các cột InputRecordMaster file
        public static string sheetName = "Master_RLMU_Receive_Final";

        public const int Ver_Col = 4;
        public const int REV_Col = 5;
        public const int Purpose_Col = 6;
        public const int Status_Col = 7;
        public const int Date_Col = 9;
        public const int TimesCount_Col = 10;
        public const int Package_Col = 11;
        public const int Discipline_Col = 12;
        public const int RedlineMarkup_Col = 13;        // Master Drawing (Receive)
        public const int PackageStampStatus_Col = 14;   // Master Drawing (Receive)

        public const int AtlasComment_Col = 18;
        public const int UpdateStatus_Col = 19;
        public const int CheckDocTitle_Col = 20;
    }
    public static class InputRecordMasterExcelData_Drawing_Submit_Final
    {
        // * Thông tin các cột InputRecordMaster file
        public static string sheetName = "Master_DWG_Submit_Final";

        public const int Ver_Col = 4;
        public const int REV_Col = 5;
        public const int Purpose_Col = 6;
        public const int Status_Col = 7;
        public const int Date_Col = 9;
        public const int TimesCount_Col = 10;
        public const int Package_Col = 11;
        public const int Discipline_Col = 12;
        public const int ModelName_Col = 13;            // Master Drawing (Submit/Feedback)

        public const int AtlasComment_Col = 18;
        public const int UpdateStatus_Col = 19;
        public const int CheckDocTitle_Col = 20;
    }
    public static class InputRecordMasterExcelData_Drawing_RFI_Final
    {
        // * Thông tin các cột InputRecordMaster file
        public static string sheetName = "Master_DWG_RFI_Final";

        public const int DrawingNo_Col = 4;
        public const int DateRequest_Col = 5;
        public const int FeedbackDate_Col = 6;
        public const int RFIStatus_Col = 7;
        public const int TimesCount_Col = 8;
    }
    public static class InputRecordMasterExcelData_Drawing_Feedback_Final
    {
        // * Thông tin các cột InputRecordMaster file
        public static string sheetName = "Master_DWG_Feedback_Final";

        public const int Ver_Col = 4;
        public const int REV_Col = 5;
        public const int Purpose_Col = 6;
        public const int Status_Col = 7;
        public const int Date_Col = 9;
        public const int TimesCount_Col = 10;
        public const int Package_Col = 11;
        public const int Discipline_Col = 12;
        public const int ModelName_Col = 13;            // Master Drawing (Submit/Feedback)

        public const int AtlasComment_Col = 18;
        public const int UpdateStatus_Col = 19;
        public const int CheckDocTitle_Col = 20;
    }

    // * MODEL
    public static class InputRecordMasterExcelData_Model_Receive
    {
        // * Thông tin các cột InputRecordMaster file
        public static string sheetName = "Master_Model_Receive";

        public const int Ver_Col = 4;
        public const int REV_Col = 5;
        public const int Purpose_Col = 6;
        public const int Status_Col = 7;
        public const int Date_Col = 9;
        public const int TimesCount_Col = 10;
        public const int Package_Col = 11;
        public const int Discipline_Col = 12;
        public const int NativeFileType_Col = 13;
        public const int SubmittedFileType_Col = 14;    // Master Model (Receive/Submit/Feedback)

        public const int AtlasComment_Col = 18;
        public const int UpdateStatus_Col = 19;
        public const int CheckDocTitle_Col = 20;
    }
    public static class InputRecordMasterExcelData_Model_Submit
    {
        // * Thông tin các cột InputRecordMaster file
        public static string sheetName = "Master_Model_Submit";

        public const int Ver_Col = 4;
        public const int REV_Col = 5;
        public const int Purpose_Col = 6;
        public const int Status_Col = 7;
        public const int Date_Col = 9;
        public const int TimesCount_Col = 10;
        public const int Package_Col = 11;
        public const int Discipline_Col = 12;
        public const int NativeFileType_Col = 13;       // Master Model (Receive/Submit/Feedback)
        public const int SubmittedFileType_Col = 14;    // Master Model (Submit/Feedback)

        public const int AtlasComment_Col = 18;
        public const int UpdateStatus_Col = 19;
        public const int CheckDocTitle_Col = 20;
    }
    public static class InputRecordMasterExcelData_Model_Feedback
    {
        // * Thông tin các cột InputRecordMaster file
        public static string sheetName = "Master_Model_Feedback";

        public const int Ver_Col = 4;
        public const int REV_Col = 5;
        public const int Purpose_Col = 6;
        public const int Status_Col = 7;
        public const int Date_Col = 9;
        public const int TimesCount_Col = 10;
        public const int Package_Col = 11;
        public const int Discipline_Col = 12;
        public const int ModelName_Col = 13;
        public const int NativeFileType_Col = 13;       // Master Model (Receive/Submit/Feedback)
        public const int SubmittedFileType_Col = 14;    // Master Model (Submit/Feedback)

        public const int AtlasComment_Col = 18;
        public const int UpdateStatus_Col = 19;
        public const int CheckDocTitle_Col = 20;
    }

    //==========================================================================
    // *** MASTER FILE - DRAWING SHEET
    public static class MasterExcelData_Drawing
    {
        // * Unique code vị trí các cột điền value
        private const string UniqueCode_Package_Col = "package#";
        private const string UniqueCode_Discipline_Col = "disc#";
        private const string UniqueCode_AllianceNo_Col = "alliance#";
        private const string UniqueCode_DocTitle_Col = "title#";
        private const string UniqueCode_WPR_Col = "wpr#";

        // * Thông tin các cột Master file
        public const string SheetName = "DRAWINGS";     // Tên Sheet sẽ ghi dữ liệu trong Master file
        public const int Start_Row = 6;                 // Dòng bắt đầu ghi dữ liệu trong Master file

        public static string sheetName = MasterExcelData_Drawing.SheetName;
        public static int Package_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_Package_Col);
        public static int Discipline_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_Discipline_Col);
        public static int Alliance_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_AllianceNo_Col);
        public static int DocTitle_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_DocTitle_Col);
        public static int WPR_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_WPR_Col);
    }

    // * WORKING
    public static class MasterExcelData_Drawing_WorkingReceive
    {
        private const string UniqueCode_TRANSID_Col = "drgworkingreceive#transid";
        private const string UniqueCode_REV_Col = "drgworkingreceive#rev";
        private const string UniqueCode_Date_Receive_Col = "drgworkingreceive#datereceive";
        private const string UniqueCode_Purpose_Col = "drgworkingreceive#purpose";
        private const string UniqueCode_Date_Issue_Col = "drgworkingreceive#dateissue";
        private const string UniqueCode_Status_Col = "drgworkingreceive#status";
        private const string UniqueCode_AtlasComment_Col = "drgworkingreceive#atlcomment";

        public static string sheetName = MasterExcelData_Drawing.SheetName;

        public static int TRANSID_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_TRANSID_Col);
        public static int REV_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_REV_Col);
        public static int Date_Receive_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_Date_Receive_Col);
        public static int Purpose_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_Purpose_Col);
        public static int Date_Issue_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_Date_Issue_Col);
        public static int Status_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_Status_Col);
        public static int AtlasComment_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_AtlasComment_Col);
    }

    // * RECEIVE
    public static class MasterExcelData_Drawing_FirstReceive
    {
        // * Unique code vị trí các cột điền value
        private const string UniqueCode_TimesCount_Col = "drgreceive1st#times";
        private const string UniqueCode_TRANSID_Col = "drgreceive1st#transid";
        private const string UniqueCode_REV_Col = "drgreceive1st#rev";
        private const string UniqueCode_Ver_Col = "drgreceive1st#ver";
        private const string UniqueCode_Purpose_Col = "drgreceive1st#purpose";
        private const string UniqueCode_Date_Col = "drgreceive1st#date";
        private const string UniqueCode_Status_Col = "drgreceive1st#status";
        private const string UniqueCode_AtlasComment_Col = "drgreceive1st#atlcomment";
        private const string UniqueCode_RedlineMarkup_Col = "drgreceive1st#redlinemarkup";
        private const string UniqueCode_PackageStampStatus_Col = "drgreceive1st#packagestampstatus";

        // * Thông tin các cột Master file
        public static string sheetName = MasterExcelData_Drawing.SheetName;

        public static int TimesCount_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_TimesCount_Col);
        public static int TRANSID_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_TRANSID_Col);
        public static int REV_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_REV_Col);
        public static int Ver_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_Ver_Col);
        public static int Purpose_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_Purpose_Col);
        public static int Date_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_Date_Col);
        public static int Status_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_Status_Col);
        public static int AtlasComment_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_AtlasComment_Col);
        public static int RedlineMarkup_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_RedlineMarkup_Col);
        public static int PackageStampStatus_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_PackageStampStatus_Col);
    }
    public static class MasterExcelData_Drawing_FinalReceive
    {
        // * Unique code vị trí các cột điền value
        private const string UniqueCode_TimesCount_Col = "drgreceivefinal#times";
        private const string UniqueCode_TRANSID_Col = "drgreceivefinal#transid";
        private const string UniqueCode_REV_Col = "drgreceivefinal#rev";
        private const string UniqueCode_Ver_Col = "drgreceivefinal#ver";
        private const string UniqueCode_Purpose_Col = "drgreceivefinal#purpose";
        private const string UniqueCode_Date_Col = "drgreceivefinal#date";
        private const string UniqueCode_Status_Col = "drgreceivefinal#status";
        private const string UniqueCode_AtlasComment_Col = "drgreceivefinal#atlcomment";
        private const string UniqueCode_RedlineMarkup_Col = "drgreceivefinal#redlinemarkup";
        private const string UniqueCode_PackageStampStatus_Col = "drgreceivefinal#packagestampstatus";

        // * Thông tin các cột Master file
        public static string sheetName = MasterExcelData_Drawing.SheetName;

        public static int TimesCount_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_TimesCount_Col);
        public static int TRANSID_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_TRANSID_Col);
        public static int REV_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_REV_Col);
        public static int Ver_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_Ver_Col);
        public static int Purpose_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_Purpose_Col);
        public static int Date_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_Date_Col);
        public static int Status_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_Status_Col);
        public static int AtlasComment_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_AtlasComment_Col);
        public static int RedlineMarkup_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_RedlineMarkup_Col);
        public static int PackageStampStatus_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_PackageStampStatus_Col);
    }

    // * SUBMIT
    public static class MasterExcelData_Drawing_FirstSubmission
    {
        // * Unique code vị trí các cột điền value
        private const string UniqueCode_TimesCount_Col = "drgsub1st#times";
        private const string UniqueCode_TRANSID_Col = "drgsub1st#transid";
        private const string UniqueCode_REV_Col = "drgsub1st#rev";
        private const string UniqueCode_Purpose_Col = "drgsub1st#purpose";
        private const string UniqueCode_Date_Col = "drgsub1st#date";
        private const string UniqueCode_Status_Col = "drgsub1st#status";
        private const string UniqueCode_AtlasComment_Col = "drgsub1st#atlcomment";

        // * Thông tin các cột Master file
        public static string sheetName = MasterExcelData_Drawing.SheetName;

        public static int TimesCount_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_TimesCount_Col);
        public static int TRANSID_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_TRANSID_Col);
        public static int REV_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_REV_Col);
        public static int Purpose_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_Purpose_Col);
        public static int Date_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_Date_Col);
        public static int Status_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_Status_Col);
        public static int AtlasComment_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_AtlasComment_Col);
    }
    public static class MasterExcelData_Drawing_FinalSubmission
    {
        // * Unique code vị trí các cột điền value
        private const string UniqueCode_TimesCount_Col = "drgsubfinal#times";
        private const string UniqueCode_TRANSID_Col = "drgsubfinal#transid";
        private const string UniqueCode_REV_Col = "drgsubfinal#rev";
        private const string UniqueCode_Purpose_Col = "drgsubfinal#purpose";
        private const string UniqueCode_Date_Col = "drgsubfinal#date";
        private const string UniqueCode_Status_Col = "drgsubfinal#status";
        private const string UniqueCode_AtlasComment_Col = "drgsubfinal#atlcomment";

        // * Thông tin các cột Master file
        public static string sheetName = MasterExcelData_Drawing.SheetName;

        public static int TimesCount_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_TimesCount_Col);
        public static int TRANSID_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_TRANSID_Col);
        public static int REV_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_REV_Col);
        public static int Purpose_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_Purpose_Col);
        public static int Date_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_Date_Col);
        public static int Status_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_Status_Col);
        public static int AtlasComment_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_AtlasComment_Col);
    }

    // * RFI
    public static class MasterExcelData_Drawing_FirstRFI    // cần sửa lại
    {
        // * Unique code vị trí các cột điền value
        private const string UniqueCode_TimesCount_Col = "drgrfi1st#times";
        private const string UniqueCode_No_Col = "drgrfi1st#no";
        private const string UniqueCode_Date_Col = "drgrfi1st#date";
        private const string UniqueCode_DateFeedback_Col = "drgrfi1st#datefeedback";
        private const string UniqueCode_Status_Col = "drgrfi1st#status";
        private const string UniqueCode_Note_Col = "drgrfi1st#note";

        // * Thông tin các cột Master file
        public static string sheetName = MasterExcelData_Drawing.SheetName;

        public static int TimesCount_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_TimesCount_Col);
        public static int No_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_No_Col);
        public static int Date_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_Date_Col);
        public static int DateFeedback_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_DateFeedback_Col);
        public static int Status_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_Status_Col);
        public static int Note_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_Note_Col);
    }
    public static class MasterExcelData_Drawing_FinalRFI
    {
        // * Unique code vị trí các cột điền value
        private const string UniqueCode_TimesCount_Col = "drgrfifinal#times";
        private const string UniqueCode_No_Col = "drgrfifinal#no";
        private const string UniqueCode_Date_Col = "drgrfifinal#date";
        private const string UniqueCode_DateFeedback_Col = "drgrfifinal#datefeedback";
        private const string UniqueCode_Status_Col = "drgrfifinal#status";
        private const string UniqueCode_Note_Col = "drgrfifinal#note";

        // * Thông tin các cột Master file
        public static string sheetName = MasterExcelData_Drawing.SheetName;

        public static int TimesCount_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_TimesCount_Col);
        public static int No_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_No_Col);
        public static int Date_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_Date_Col);
        public static int DateFeedback_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_DateFeedback_Col);
        public static int Status_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_Status_Col);
        public static int Note_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_Note_Col);
    }

    // * FEEDBACK
    public static class MasterExcelData_Drawing_FirstFeedback
    {
        // * Unique code vị trí các cột điền value
        private const string UniqueCode_TimesCount_Col = "drgfeedback1st#times";
        private const string UniqueCode_TRANSID_Col = "drgfeedback1st#transid";
        private const string UniqueCode_Status_Col = "drgfeedback1st#status";
        private const string UniqueCode_Date_Col = "drgfeedback1st#date";
        private const string UniqueCode_AtlasComment_Col = "drgfeedback1st#atlcomment";

        // * Thông tin các cột Master file
        public static string sheetName = MasterExcelData_Drawing.SheetName;

        public static int TimesCount_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_TimesCount_Col);
        public static int TRANSID_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_TRANSID_Col);
        public static int Status_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_Status_Col);
        public static int Date_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_Date_Col);
        public static int AtlasComment_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_AtlasComment_Col);
    }
    public static class MasterExcelData_Drawing_FinalFeedback
    {
        // * Unique code vị trí các cột điền value
        private const string UniqueCode_TimesCount_Col = "drgfeedbackfinal#times";
        private const string UniqueCode_TRANSID_Col = "drgfeedbackfinal#transid";
        private const string UniqueCode_Status_Col = "drgfeedbackfinal#status";
        private const string UniqueCode_Date_Col = "drgfeedbackfinal#date";
        private const string UniqueCode_AtlasComment_Col = "drgfeedbackfinal#atlcomment";

        // * Thông tin các cột Master file
        public static string sheetName = MasterExcelData_Drawing.SheetName;

        public static int TimesCount_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_TimesCount_Col);
        public static int TRANSID_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_TRANSID_Col);
        public static int Status_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_Status_Col);
        public static int Date_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_Date_Col);
        public static int AtlasComment_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_AtlasComment_Col);
    }

    //==========================================================================
    // *** MASTER FILE - MODEL SHEET
    public static class MasterExcelData_Model
    {
        // * Unique code vị trí các cột điền value
        private const string UniqueCode_Package_Col = "package#";
        private const string UniqueCode_Discipline_Col = "disc#";
        private const string UniqueCode_AllianceNo_Col = "alliance#";
        private const string UniqueCode_DocTitle_Col = "title#";
        private const string UniqueCode_WPR_Col = "wpr#";

        // * Thông tin các cột Master file
        public const string SheetName = "MASTER MODEL LIST";    // Tên Sheet sẽ ghi dữ liệu trong Master file
        public const int Start_Row = 14;                        // Dòng bắt đầu ghi dữ liệu trong Master file

        public static string sheetName = SheetName;
        public static int Package_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_Package_Col);
        public static int Discipline_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_Discipline_Col);
        public static int Alliance_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_AllianceNo_Col);
        public static int DocTitle_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_DocTitle_Col);
        public static int WPR_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_WPR_Col);


    }

    public static class MasterExcelData_Model_Receive
    {
        // * Unique code vị trí các cột điền value
        private const string UniqueCode_TimesCount_Col = "modelreceive#times";
        private const string UniqueCode_TRANSID_Col = "modelreceive#transid";
        private const string UniqueCode_REV_Col = "modelreceive#rev";
        private const string UniqueCode_Ver_Col = "modelreceive#ver";
        private const string UniqueCode_Purpose_Col = "modelreceive#purpose";
        private const string UniqueCode_Date_Col = "modelreceive#date";
        private const string UniqueCode_Status_Col = "modelreceive#status";
        private const string UniqueCode_AtlasComment_Col = "modelreceive#atlcomment";
        private const string UniqueCode_FileType_Col = "modelreceive#filetype";

        // * Thông tin các cột Master file
        public static string sheetName = MasterExcelData_Model.SheetName;

        public static int TimesCount_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_TimesCount_Col);
        public static int TRANSID_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_TRANSID_Col);
        public static int REV_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_REV_Col);
        public static int Ver_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_Ver_Col);
        public static int Purpose_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_Purpose_Col);
        public static int Date_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_Date_Col);
        public static int Status_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_Status_Col);
        public static int AtlasComment_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_AtlasComment_Col);
        public static int FileType_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_FileType_Col);
    }
    public static class MasterExcelData_Model_Submission
    {
        // * Unique code vị trí các cột điền value
        private const string UniqueCode_TimesCount_Col = "modelsub#times";
        private const string UniqueCode_TRANSID_Col = "modelsub#transid";
        private const string UniqueCode_REV_Col = "modelsub#rev";
        private const string UniqueCode_Purpose_Col = "modelsub#purpose";
        private const string UniqueCode_Date_Col = "modelsub#date";
        private const string UniqueCode_Status_Col = "modelsub#status";
        private const string UniqueCode_AtlasComment_Col = "modelsub#atlcomment";
        private const string UniqueCode_FileType_Col = "modelsub#filetype";

        // * Thông tin các cột Master file
        public static string sheetName = MasterExcelData_Model.SheetName;

        public static int TimesCount_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_TimesCount_Col);
        public static int TRANSID_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_TRANSID_Col);
        public static int REV_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_REV_Col);
        public static int Purpose_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_Purpose_Col);
        public static int Date_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_Date_Col);
        public static int Status_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_Status_Col);
        public static int AtlasComment_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_AtlasComment_Col);
        public static int FileType_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_FileType_Col);
    }
    public static class MasterExcelData_Model_Feedback
    {
        // * Unique code vị trí các cột điền value
        private const string UniqueCode_TimesCount_Col = "modelfeedback#times";
        private const string UniqueCode_TRANSID_Col = "modelfeedback#transid";
        private const string UniqueCode_Status_Col = "modelfeedback#status";
        private const string UniqueCode_Date_Col = "modelfeedback#date";
        private const string UniqueCode_AtlasComment_Col = "modelfeedback#atlcomment";

        // * Thông tin các cột Master file
        public static string sheetName = MasterExcelData_Model.SheetName;

        public static int TimesCount_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_TimesCount_Col);
        public static int TRANSID_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_TRANSID_Col);
        public static int Status_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_Status_Col);
        public static int Date_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_Date_Col);
        public static int AtlasComment_Col = GetColumnIndexByUniqueCode(TrackingInputData.txtFilePath_Master, sheetName, UniqueCode_AtlasComment_Col);
    }



}