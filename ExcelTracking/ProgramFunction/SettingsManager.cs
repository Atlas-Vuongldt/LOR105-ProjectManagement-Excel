using System;
using System.IO;
using System.Windows.Forms;
using Newtonsoft.Json;

namespace SettingsManager
{
    // 📋 FormSettings cho trường hợp 1
    public class FormSettings_GetInputData
    {
        public string InputFolder { get; set; } = "";
        public string OutputFolder { get; set; } = "";
        public string MasterFile { get; set; } = "";
    }

    // 📋 FormSettings cho trường hợp 2
    public class FormSettings_MainTracking
    {
        public string MasterFile { get; set; } = "";
        public string InputDataFile { get; set; } = "";
        public string RecordMasterFile { get; set; } = "";
        // Có thể thêm properties khác cho trường hợp 2
        // public string AdditionalPath { get; set; } = "";
    }

    public static class SettingsManagerConfig
    {
        // 📁 Đường dẫn cho settings 1
        private static string GetSettingsPath_GetInputData()
        {
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            return Path.Combine(desktopPath, "FormSettings_GetInputData.atl");
        }

        // 📁 Đường dẫn cho settings 2  
        private static string GetSettingsPath_MainTracking()
        {
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            return Path.Combine(desktopPath, "FormSettings_MainTracking.atl");
        }

        // 📂 Load Settings 1
        public static FormSettings_GetInputData LoadSettings_GetInputData()
        {
            try
            {
                string filePath = GetSettingsPath_GetInputData();

                if (!File.Exists(filePath))
                {
                    // 🆕 Tạo settings mặc định và save luôn để lần sau có file
                    var defaultSettings = new FormSettings_GetInputData();
                    SaveSettings_GetInputData(defaultSettings);
                    return defaultSettings;
                }

                string json = File.ReadAllText(filePath);
                return JsonConvert.DeserializeObject<FormSettings_GetInputData>(json) ?? new FormSettings_GetInputData();
            }
            catch (Exception ex)
            {
                return new FormSettings_GetInputData();
            }
        }

        // 📂 Load Settings 2
        public static FormSettings_MainTracking LoadSettings_MainTracking()
        {
            try
            {
                string filePath = GetSettingsPath_MainTracking();

                if (!File.Exists(filePath))
                {
                    // 🆕 Tạo settings mặc định và save luôn để lần sau có file
                    var defaultSettings = new FormSettings_MainTracking();
                    SaveSettings_MainTracking(defaultSettings);
                    return defaultSettings;
                }

                string json = File.ReadAllText(filePath);
                return JsonConvert.DeserializeObject<FormSettings_MainTracking>(json) ?? new FormSettings_MainTracking();
            }
            catch (Exception ex)
            {
                return new FormSettings_MainTracking();
            }
        }

        // 💾 Save Settings 1
        public static void SaveSettings_GetInputData(FormSettings_GetInputData settings)
        {
            try
            {
                string json = JsonConvert.SerializeObject(settings, Formatting.Indented);
                File.WriteAllText(GetSettingsPath_GetInputData(), json);
            }
            catch (Exception ex)
            {
                // Handle error
            }
        }

        // 💾 Save Settings 2
        public static void SaveSettings_MainTracking(FormSettings_MainTracking settings)
        {
            try
            {
                string json = JsonConvert.SerializeObject(settings, Formatting.Indented);
                File.WriteAllText(GetSettingsPath_MainTracking(), json);
            }
            catch (Exception ex)
            {
                // Handle error
            }
        }
    }
}