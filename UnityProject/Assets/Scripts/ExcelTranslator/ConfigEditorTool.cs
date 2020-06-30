#if UNITY_EDITOR
using UnityEditor;
#endif
using UnityEngine;
using System.Collections.Generic;
using System.IO;
using System.Text;
using OfficeOpenXml;
using Engine.Core.ExcelTranslator;

namespace Engine.Config
{
    public class DataEditorTool
    {
        private readonly static string dataPath = Application.dataPath;
        public readonly static string ExcelFloder = Application.dataPath + "/../../ExcelFiles/";
        public readonly static string ByteFilePath = dataPath + "/Scripts/Sample/GenBytes/";
        public readonly static string JsonFilePath = dataPath + "/Scripts/Sample/GenJsonConfig/";
        public readonly static string LuaDataEntryPath = dataPath + "/Scripts/Sample/GenLuaTable/";
        public readonly static string CSharpDataEntryPath = dataPath + "/Scripts/Sample/GenCSharpClass/";

#if UNITY_EDITOR
        [MenuItem("Tools/Config工具/一键生成配置(Lua & Byte & C#)", false)]
        private static void TranslatorExcelConfigs()
        {
            TranslatorExcelConfigs(true);
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

                //json
                string json_path = Path.Combine(JsonFilePath, translator.sheetName + ".txt");
                File.WriteAllBytes(json_path, UTF8Encoding.UTF8.GetBytes(translator.ToJson()));

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
#endif
    }
}