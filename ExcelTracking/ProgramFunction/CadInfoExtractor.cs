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
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Exception = Autodesk.AutoCAD.Runtime.Exception;
#endregion
namespace ExcelTracking
{
    // 📋 Model Classes
    public class CadBlockConfig
    {
        public List<BlockConfig> Blocks { get; set; } = new List<BlockConfig>();
    }

    public class BlockConfig
    {
        public string Name { get; set; }
        public int MaxRevisions { get; set; }
        public PatternConfig Patterns { get; set; }
        public string MainRevTag { get; set; }
        public int RevStartIndex { get; set; } = 1;
        public bool UseSpecialFirstRev { get; set; } = false;
    }

    public class PatternConfig
    {
        public string Rev { get; set; }
        public string Date { get; set; }
        public string Amendment { get; set; }
    }

    public class CadInfoExtractor
    {
        public static string settingFolder = @"\\SRVPRD4\Structure\ATLAS TOOLS - STR\LOR105\Setting Files";

        public static void GetCadInfo(string fileName,
            out string Out_maxRevValue,
            out string Out_maxRevDate,
            out string Out_maxRevAmendment,
            out string Out_Status,
            string configPath = "CadBlockConfig.atl")
        {
            Out_maxRevValue = "";
            Out_maxRevDate = "";
            Out_maxRevAmendment = "";
            Out_Status = "";

            configPath = Path.Combine(settingFolder, "CadBlockConfig.atl");

            // 📖 Load configuration từ JSON with comments
            var config = LoadConfiguration(configPath);
            if (config == null || !config.Blocks.Any())
            {
                Out_Status = "Configuration file not found or invalid";
                return;
            }

            using (var db = new Database(false, true))
            {
                if (!fileName.EndsWith("dwg")) { return; }

                try
                {
                    db.ReadDwgFile(fileName, FileOpenMode.OpenForReadAndAllShare, false, null);
                }
                catch (Autodesk.AutoCAD.Runtime.Exception e)
                {
                    Out_Status = $"Error reading DWG file: {e.Message}";
                    return;
                }

                using (Transaction trans = db.TransactionManager.StartTransaction())
                {
                    try
                    {
                        BlockTable bt = trans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                        DBDictionary layouts = (DBDictionary)trans.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);

                        // 🎯 Biến lưu trữ REV lớn nhất
                        int maxRevNumber = -1;
                        string maxRevValue = "";
                        string maxRevDate = "";
                        string maxRevAmendment = "";
                        string mainRevValue = "";

                        foreach (DBDictionaryEntry entry in layouts)
                        {
                            Layout layout = (Layout)trans.GetObject(entry.Value, OpenMode.ForRead);

                            if (layout.TabOrder == 1) // Tab Model
                            {
                                var btrL = (BlockTableRecord)trans.GetObject(layout.BlockTableRecordId, OpenMode.ForRead);

                                foreach (ObjectId entId in btrL)
                                {
                                    var entity = (Entity)trans.GetObject(entId, OpenMode.ForRead);

                                    if (entity is BlockReference blockReference)
                                    {
                                        BlockTableRecord block = trans.GetObject(bt[blockReference.Name], OpenMode.ForRead) as BlockTableRecord;

                                        if (!block.IsFromExternalReference)
                                        {
                                            // 🔍 Tìm block config tương ứng
                                            var blockConfig = config.Blocks.FirstOrDefault(b => b.Name == block.Name);
                                            if (blockConfig != null)
                                            {
                                                ProcessBlock(blockReference, blockConfig, trans, ref maxRevNumber,
                                                           ref maxRevValue, ref maxRevDate, ref maxRevAmendment, ref mainRevValue);
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        // ✅ Kiểm tra trạng thái
                        if (!string.IsNullOrEmpty(maxRevValue) && !string.IsNullOrEmpty(mainRevValue))
                        {
                            if (maxRevValue != mainRevValue)
                            {
                                Out_Status = "REV in CAD file Not Match";
                            }
                        }

                        Out_maxRevValue = maxRevValue;
                        Out_maxRevDate = maxRevDate;
                        Out_maxRevAmendment = maxRevAmendment;
                    }
                    catch (Exception ex)
                    {
                        Out_Status = $"Error processing CAD data: {ex.Message}";
                    }
                    trans.Commit();
                }
            }
        }

        // 🚀 Load configuration with comment support
        private static CadBlockConfig LoadConfiguration(string configPath)
        {
            try
            {
                if (!File.Exists(configPath))
                {
                    var defaultConfig = CreateDefaultConfig();
                    SaveConfiguration(defaultConfig, configPath);
                    return defaultConfig;
                }

                string jsonContent = File.ReadAllText(configPath);

                // 🔥 Loại bỏ comments trước khi parse
                jsonContent = RemoveJsonComments(jsonContent);

                var config = JsonConvert.DeserializeObject<CadBlockConfig>(jsonContent);

                if (config?.Blocks == null || !config.Blocks.Any())
                {
                    return CreateDefaultConfig();
                }

                return config;
            }
            catch (Exception ex)
            {
                // Log error nếu cần
                Console.WriteLine($"Config load error: {ex.Message}");
                return CreateDefaultConfig();
            }
        }

        // 🧹 Remove JSON comments (// style)
        private static string RemoveJsonComments(string json)
        {
            var lines = json.Split(new[] { '\r', '\n' }, StringSplitOptions.None);
            var cleanedLines = new List<string>();

            foreach (var line in lines)
            {
                // Tìm vị trí comment // (không nằm trong string)
                int commentIndex = FindCommentIndex(line);

                if (commentIndex >= 0)
                {
                    // Lấy phần trước comment, trim whitespace
                    string cleanLine = line.Substring(0, commentIndex).TrimEnd();
                    cleanedLines.Add(cleanLine);
                }
                else
                {
                    cleanedLines.Add(line);
                }
            }

            return string.Join("\n", cleanedLines);
        }

        // 🎯 Tìm vị trí comment, bỏ qua nếu nằm trong string
        private static int FindCommentIndex(string line)
        {
            bool inString = false;
            bool escapeNext = false;

            for (int i = 0; i < line.Length - 1; i++)
            {
                char current = line[i];
                char next = line[i + 1];

                if (escapeNext)
                {
                    escapeNext = false;
                    continue;
                }

                if (current == '\\')
                {
                    escapeNext = true;
                    continue;
                }

                if (current == '"')
                {
                    inString = !inString;
                    continue;
                }

                // Nếu không trong string và gặp //
                if (!inString && current == '/' && next == '/')
                {
                    return i;
                }
            }

            return -1; // Không tìm thấy comment
        }

        // ⚙️ Process individual block
        private static void ProcessBlock(BlockReference blockReference, BlockConfig blockConfig,
            Transaction trans, ref int maxRevNumber, ref string maxRevValue,
            ref string maxRevDate, ref string maxRevAmendment, ref string mainRevValue)
        {
            // 📋 Lấy tất cả attributes
            Dictionary<string, string> attValues = new Dictionary<string, string>();
            foreach (ObjectId id in blockReference.AttributeCollection)
            {
                AttributeReference attRef = (AttributeReference)trans.GetObject(id, OpenMode.ForWrite);
                attValues[attRef.Tag] = attRef.TextString;
            }

            // 🔎 Tìm REV lớn nhất
            for (int rev = blockConfig.RevStartIndex; rev <= blockConfig.MaxRevisions; rev++)
            {
                string revTag = GetRevTag(blockConfig, rev);
                string dateTag = string.Format(blockConfig.Patterns.Date, rev);
                string amendmentTag = string.Format(blockConfig.Patterns.Amendment, rev);

                if (attValues.ContainsKey(revTag) && !string.IsNullOrWhiteSpace(attValues[revTag]))
                {
                    if (rev > maxRevNumber)
                    {
                        maxRevNumber = rev;
                        maxRevValue = attValues[revTag];
                        maxRevDate = attValues.ContainsKey(dateTag) ? attValues[dateTag] : "";
                        maxRevAmendment = attValues.ContainsKey(amendmentTag) ? attValues[amendmentTag] : "";
                    }
                }
            }

            // 🎯 Lấy main REV
            if (!string.IsNullOrEmpty(blockConfig.MainRevTag))
            {
                if (attValues.ContainsKey(blockConfig.MainRevTag))
                {
                    mainRevValue = attValues[blockConfig.MainRevTag];
                }
            }
            else
            {
                mainRevValue = maxRevValue; // Không có main REV thì lấy max REV
            }
        }

        // 🏷️ Get revision tag based on config
        private static string GetRevTag(BlockConfig blockConfig, int revNumber)
        {
            if (blockConfig.UseSpecialFirstRev && revNumber == blockConfig.RevStartIndex)
            {
                return blockConfig.MainRevTag ?? "REV"; // Fallback nếu MainRevTag null
            }
            return string.Format(blockConfig.Patterns.Rev, revNumber);
        }

        // 🏭 Create default configuration
        private static CadBlockConfig CreateDefaultConfig()
        {
            return new CadBlockConfig
            {
                Blocks = new List<BlockConfig>
            {
                new BlockConfig
                {
                    Name = "PTA_A1",
                    MaxRevisions = 7,
                    Patterns = new PatternConfig
                    {
                        Rev = "REV{0}",
                        Date = "DATE{0}",
                        Amendment = "AMENDMENT{0}"
                    },
                    MainRevTag = "REV",
                    RevStartIndex = 1,
                    UseSpecialFirstRev = true
                },
                new BlockConfig
                {
                    Name = "MRWA_A1_CON_HTBLK",
                    MaxRevisions = 9,
                    Patterns = new PatternConfig
                    {
                        Rev = "REV-{0:00}",
                        Date = "REV-{0:00}-DATE",
                        Amendment = "REV-{0:00}-DESC"
                    },
                    MainRevTag = null,
                    RevStartIndex = 1,
                    UseSpecialFirstRev = false
                }
            }
            };
        }

        // 💾 Save configuration to file
        private static void SaveConfiguration(CadBlockConfig config, string configPath)
        {
            try
            {
                string json = JsonConvert.SerializeObject(config, Formatting.Indented);

                // Thêm header comment vào file
                string header = @"// CAD Block Configuration File
                    // Supports // style comments
                    // Patterns use C# string.Format syntax
                    // {0} will be replaced with revision number
                    ";
                File.WriteAllText(configPath, header + json);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving config: {ex.Message}");
            }
        }

        public static void OpenFolder(string folderPath)
        {
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


        // 🧪 Test method
        [CommandMethod("TESTINGCADBLOCK")]
        public static void TestConfiguration(string configPath = "CadBlockConfig.atl")
        {
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            configPath = Path.Combine(desktopPath, "CadBlockConfig.atl");

            try
            {
                var config = LoadConfiguration(configPath);

                string message = "🔧 CAD BLOCK CONFIGURATION TEST\n";
                message += "═══════════════════════════════════\n\n";
                message += $"📂 Config file: {Path.GetFileName(configPath)}\n";
                message += $"📊 Loaded blocks: {config.Blocks.Count}\n\n";

                message += "📋 BLOCK DETAILS:\n";
                message += "─────────────────────────\n";

                foreach (var block in config.Blocks)
                {
                    message += $"🔹 {block.Name}\n";
                    message += $"   • Max revisions: {block.MaxRevisions}\n";
                    message += $"   • Pattern: {block.Patterns.Rev}\n";
                    message += $"   • Special first: {(block.UseSpecialFirstRev ? "Yes" : "No")}\n\n";
                }

                MessageBox.Show(message, "CAD Configuration Test Result",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ ERROR LOADING CONFIG:\n\n{ex.Message}",
                    "Configuration Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }














        // ============================================================================================================
        // *** AUTOCAD
        #region

        public static void GetCadInfo_OLD(string fileName, 
            out string Out_maxRevValue, 
            out string Out_maxRevDate, 
            out string Out_maxRevAmendment,
            out string Out_Status)
        {
            Out_maxRevValue = "";
            Out_maxRevDate = "";
            Out_maxRevAmendment = "";
            Out_Status = "";

            using (var db = new Database(false, true))
            {
                if (!fileName.EndsWith("dwg")) { return; }
                try
                {
                    db.ReadDwgFile(fileName, FileOpenMode.OpenForReadAndAllShare, false, null);
                }
                catch (Autodesk.AutoCAD.Runtime.Exception e)
                {
                    //MessageBox.Show(fileName + "\n" + e.Message + "\n" + e.ToString(), "Error",
                    //    MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                using (Transaction trans = db.TransactionManager.StartTransaction())
                {
                    try
                    {
                        BlockTable bt = trans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                        DBDictionary layouts = (DBDictionary)trans.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);

                        // Biến lưu trữ REV lớn nhất và các thông tin liên quan
                        int maxRevNumber = -1;
                        string maxRevValue = "";
                        string maxRevDate = "";
                        string maxRevAmendment = "";
                        string mainRevValue = ""; // Lưu giá trị của REV chính (att REV)

                        foreach (DBDictionaryEntry entry in layouts)
                        {
                            // Get Layout
                            Layout layout = (Layout)trans.GetObject(entry.Value, OpenMode.ForRead);

                            // Vào Tab Model
                            if (layout.TabOrder == 1)
                            {
                                var btrL = (BlockTableRecord)trans.GetObject(layout.BlockTableRecordId, OpenMode.ForRead);

                                foreach (ObjectId entId in btrL)
                                {
                                    // Get Objects in Layout
                                    var entity = (Entity)trans.GetObject(entId, OpenMode.ForRead);

                                    if (entity is BlockReference)
                                    {
                                        BlockReference blockReference = entity as BlockReference;
                                        BlockTableRecord block = trans.GetObject(bt[blockReference.Name], OpenMode.ForRead) as BlockTableRecord;

                                        if (!block.IsFromExternalReference)
                                        {
                                            string blockName_01 = "PTA_A1";
                                            string blockName_02 = "MRWA_A1_CON_HTBLK";

                                            // Tên block
                                            string blockName = block.Name;
                                            if (block.Name == blockName_01 || blockName == blockName_02)
                                            {
                                                Dictionary<string, string> attValues = new Dictionary<string, string>();

                                                foreach (ObjectId id in blockReference.AttributeCollection)
                                                {
                                                    AttributeReference attRef = (AttributeReference)trans.GetObject(id, OpenMode.ForWrite);
                                                    attValues[attRef.Tag] = attRef.TextString;
                                                }

                                                // Xác định các pattern tag dựa vào tên block
                                                string revPattern = "", datePattern = "", amendmentPattern = "";
                                                int maxRevisions = 0;

                                                // Cấu hình pattern dựa trên loại block
                                                if (blockName == blockName_01)
                                                {
                                                    revPattern = "REV{0}";              // REV, REV2, REV3, ...
                                                    datePattern = "DATE{0}";            // DATE1, DATE2, DATE3, ...
                                                    amendmentPattern = "AMENDMENT{0}";  // AMENDMENT1, AMENDMENT2, ...
                                                    maxRevisions = 7;
                                                }
                                                else if (blockName == blockName_02)
                                                {
                                                    revPattern = "REV-{0:00}";              // REV-01, REV-02, ...
                                                    datePattern = "REV-{0:00}-DATE";        // REV-01-DATE, REV-02-DATE, ...
                                                    amendmentPattern = "REV-{0:00}-DESC";   // REV-01-DESC, REV-02-DESC, ...
                                                    maxRevisions = 9;
                                                }

                                                // Tìm REV lớn nhất có giá trị
                                                for (int rev = 1; rev <= maxRevisions; rev++)
                                                {
                                                    string revTag = (blockName == blockName_01 && rev == 1) ? "REV" : string.Format(revPattern, rev);
                                                    string dateTag = string.Format(datePattern, rev);
                                                    string amendmentTag = string.Format(amendmentPattern, rev);

                                                    if (attValues.ContainsKey(revTag) && !string.IsNullOrWhiteSpace(attValues[revTag]))
                                                    {
                                                        if (rev > maxRevNumber)
                                                        {
                                                            maxRevNumber = rev;
                                                            maxRevValue = attValues[revTag];
                                                            maxRevDate = attValues.ContainsKey(dateTag) ? attValues[dateTag] : "";
                                                            maxRevAmendment = attValues.ContainsKey(amendmentTag) ? attValues[amendmentTag] : "";
                                                        }
                                                    }
                                                }

                                                // Chỉ lấy main REV đối với PTA_A1
                                                if (blockName == blockName_01)
                                                {
                                                    if (attValues.ContainsKey("REV"))
                                                    {
                                                        mainRevValue = attValues["REV"];
                                                    }
                                                }

                                                // Đối với các block không phải PTA_A1, lấy REV max làm main REV
                                                if (blockName != blockName_01)
                                                {
                                                    mainRevValue = maxRevValue;
                                                }

                                                // So sánh REV lớn nhất với REV chính (chỉ kiểm tra cho PTA_A1)
                                                if (blockName == blockName_01 && !string.IsNullOrEmpty(maxRevValue) && !string.IsNullOrEmpty(mainRevValue))
                                                {
                                                    bool err = maxRevValue != mainRevValue;
                                                    if (err)
                                                    {
                                                        Out_Status = "REV in CAD file Not Match";
                                                    }
                                                }

                                                Out_maxRevValue = maxRevValue;
                                                Out_maxRevDate = maxRevDate;
                                                Out_maxRevAmendment = maxRevAmendment;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    catch { }
                    trans.Commit();
                }
            }
        }
        #endregion
    }
}