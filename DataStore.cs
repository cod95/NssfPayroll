using System;
using System.IO;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace NssfPayroll
{
    /// <summary>
    /// بيحفظ ويقرا كل البيانات من ملف JSON:
    /// %AppData%\NssfPayroll\data.json
    /// (للنسخ الاحتياطي: انسخ هالملف).
    /// </summary>
    public static class DataStore
    {
        private static readonly string Folder =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "NssfPayroll");

        private static readonly string FilePath = Path.Combine(Folder, "data.json");

        private static readonly JsonSerializerOptions Options = new JsonSerializerOptions
        {
            WriteIndented = true,
            IgnoreReadOnlyProperties = true,                              // ما منحفظ الخصائص المحسوبة
            DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull, // ما منكتب الحقول الفاضية (تواريخ اختيارية...)
            Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping         // ليضل العربي مقروء بالملف
        };

        public static string DataFilePath { get { return FilePath; } }

        public static AppData Load()
        {
            if (!File.Exists(FilePath)) return new AppData();

            try
            {
                string json = File.ReadAllText(FilePath);
                AppData data = JsonSerializer.Deserialize<AppData>(json, Options) ?? new AppData();

                // ملف من نسخة قديمة (مؤسسة وحدة، أو أجور الفصل كاملة): منحوّلو للشكل الجديد،
                // ومنحتفظ بنسخة عن الملف القديم (باسم فيه التاريخ، لحتى ما تمحي نسخة أقدم)
                if (data.MigrateLegacy())
                {
                    try
                    {
                        File.Copy(FilePath, FilePath + ".backup-" + DateTime.Now.ToString("yyyyMMdd-HHmmss"), true);
                        Save(data);
                    }
                    catch (Exception) { }
                }

                return data;
            }
            catch (Exception)
            {
                // الملف تالف: منحتفظ بنسخة منو وبنبلّش من جديد بدل ما يوقع البرنامج
                try
                {
                    File.Copy(FilePath, FilePath + ".broken-" + DateTime.Now.ToString("yyyyMMdd-HHmmss"), true);
                }
                catch (Exception) { }
                return new AppData();
            }
        }

        public static void Save(AppData data)
        {
            Directory.CreateDirectory(Folder);

            // منكتب لملف مؤقت أول وبعدين منبدّل، لحتى ما يتلف الملف إذا انقطعت الكهربا وسط الحفظ
            string tmp = FilePath + ".tmp";
            File.WriteAllText(tmp, JsonSerializer.Serialize(data, Options));
            File.Copy(tmp, FilePath, true);
            File.Delete(tmp);
        }
    }
}
