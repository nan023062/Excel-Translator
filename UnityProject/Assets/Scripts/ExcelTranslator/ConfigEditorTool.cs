#if UNITY_EDITOR
using UnityEditor;
using UnityEngine;
using System.Collections.Generic;
using System.IO;
using System.Text;
using OfficeOpenXml;
using Engine.Core.ExcelTranslator;

namespace Engine.Core
{
    public class DataEditorTool : Editor
    {
        private readonly static string dataPath = Application.dataPath;
        public readonly static string ExcelFloder = Application.dataPath + "/../../DevelopTools/ExcelConfig/";
        public readonly static string ByteFilePath = dataPath + "/GameAssets/Configs/";
        public readonly static string JsonFilePath = dataPath + "/GameAssets/Configs/Json/";
        public readonly static string LuaDataEntryPath = dataPath + "/GameAssets/LuaCode/GenConfigTables/";
        public readonly static string CSharpDataEntryPath = dataPath + "/Scripts/GlobalDataClass/GenConfigClass/";

        [MenuItem("Tools/Config工具/一键生成配置(Lua & Byte & C#)", false)]
        private static void TranslatorExcelConfigs()
        {
            TranslatorExcelConfigs(true);
        }

        [MenuItem("Tools/Config工具/Test", false)]
        public static void testc()
        {
            var excelSheet = ExcelTranslatorUtility.ReadExcelSheet(ExcelFloder, "god_damage");
            var json = TranslatorTable.ToJson(excelSheet, 0X7FFFFFFF);
        }

        public static void TranslatorExcelConfigs(bool showDialog, int readMask = 0x7FFFFFFF)
        {
            if(showDialog) EditorUtility.DisplayProgressBar("Translator Excel Configs", "读取Excel文件", 0f);

            var ExcelSheets = ReadAllExcelConfigs(showDialog);

            int count = ExcelSheets.Count;
            int index = 1;
            foreach (var excelSheet in ExcelSheets.Values)
            {
                var translator = new TranslatorTable(excelSheet, readMask);
                
                //byte
                string byte_path = Path.Combine(ByteFilePath, translator.sheetName + ".byte");
                File.WriteAllBytes(byte_path, translator.ToDataEntryBytes());

                //lua
                string lua_path = Path.Combine(LuaDataEntryPath, translator.sheetName + ".lua");
                File.WriteAllBytes(lua_path, UTF8Encoding.UTF8.GetBytes(translator.ToLuaTable()));

                //c#
                string csharp_path = Path.Combine(CSharpDataEntryPath, translator.sheetName + ".cs");
                File.WriteAllBytes(csharp_path, UTF8Encoding.UTF8.GetBytes(translator.ToDataEntryClass()));

                if (showDialog)
                {
                    float prog = index *1f / count ;
                    string content = string.Format("转换配置表:【{0}】", translator.sheetName);
                    EditorUtility.DisplayProgressBar("Translator Excel Configs", content, prog);
                    index++;
                }
            }

            if (showDialog)
            {
                EditorUtility.DisplayProgressBar("Translator Excel Configs", "配置表转换完成", 1);
                EditorUtility.ClearProgressBar();
                EditorUtility.DisplayDialog("Translator Excel Configs", "配置表转换完成！", "OK");
            }
        }

        public static Dictionary<string, ExcelWorksheet> ReadAllExcelConfigs(bool showDialog)
        {
            var ExcelSheets = ExcelTranslatorUtility.ReadALLExcelSheets(ExcelFloder, (excelName, prog) =>
            {
                if (showDialog)
                {
                    string content = string.Format("读取Excel:【{0}】 ...", excelName);
                    EditorUtility.DisplayProgressBar("读取配置表", content, 1f * prog);
                }
            });
            if (showDialog) EditorUtility.ClearProgressBar();
            return ExcelSheets;
        }

        [MenuItem("Tools/Config工具/生成Excel的映射文件", false)]
        private static void WriteExcelNameToPath()
        {
            ExcelTranslatorUtility.WriteExcelNameToPath(ExcelFloder, (excelName, prog) =>
            {
                string content = string.Format("读取Excel:【{0}】 ...", excelName);
                EditorUtility.DisplayProgressBar("生成Excel的映射文件...", content, 1f * prog);
            });
            EditorUtility.ClearProgressBar();
        }

        [MenuItem("Tools/Config工具/配置表检查/String表检查", false)]
        private static void CheckStringKeyUniqueness()
        {
            CheckExcelKeyUniqueness("string_cfg");
            CheckExcelNotEmptyCell("string_cfg");
        }

        private static void CheckExcelKeyUniqueness(string configName)
        {
            var excelSheet = ExcelTranslatorUtility.ReadExcelSheet(ExcelFloder, configName);
            var checkIDMap = new Dictionary<string, int>();

            var table = new TranslatorTable(excelSheet, 0x7FFFFFFF);
            for (int i = 0; i < table.nRow; i++)
            {
                string id = table.ID(i);
                int count = 0;
                checkIDMap.TryGetValue(id, out count);
                count++;
                checkIDMap[id] = count;
            }

            bool result = true;
            foreach (var configIDCount in checkIDMap)
            {
                if (configIDCount.Value > 1)
                {
                    Debug.LogErrorFormat("ID = {0}不唯一！，存在{1}个！", configIDCount.Key, configIDCount.Value);
                    result = false;
                }
            }
            if(result) Debug.LogFormat("配置表{0}的ID没有发现重复...", configName);
        }

        private static void CheckExcelNotEmptyCell(string configName)
        {
            var excelSheet = ExcelTranslatorUtility.ReadExcelSheet(ExcelFloder, configName);
            var table = new TranslatorTable(excelSheet, 0x7FFFFFFF);
        }
    }
}
#endif
