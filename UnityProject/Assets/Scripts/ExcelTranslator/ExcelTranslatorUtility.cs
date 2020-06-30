using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Engine.Core.ExcelTranslator
{
    public static class ExcelTranslatorUtility
    {
        public const string NAME_TO_PATH = "sheetNameToPath.txt";

        #region Excel Option Tools

        private static Dictionary<string, ExcelWorksheet> mExcelSheetCaches = null;

        public static List<ExcelWorksheet> ReadExcelSheets(string filePath)
        {
            if (File.Exists(filePath) && filePath.EndsWith(".xlsx"))
            {
                try
                {
                    FileInfo fileInfo = new FileInfo(filePath);
                    return ReadExcelSheets(fileInfo);
                }
                catch (System.Exception ex)
                {
                    UnityEngine.Debug.LogErrorFormat("Read {0} Failed! msg = {1}!",filePath, ex.Message);
                }
            }
            return null;
        }

        public static List<ExcelWorksheet> ReadExcelSheets(FileInfo excelFile)
        {
            List<ExcelWorksheet> sheetLst = null;
            try
            {
                sheetLst = new List<ExcelWorksheet>();
                ExcelPackage ep = new ExcelPackage(excelFile);
                ExcelWorksheets worksheets = ep.Workbook.Worksheets;
                for (int i = 1; i <= worksheets.Count; i++)
                {
                    ExcelWorksheet sheet = worksheets[i];
                    sheetLst.Add(sheet);
                }
            }
            catch (System.Exception ex)
            {
                throw new System.Exception(ex.Message);
            }
            return sheetLst;
        }

        public static void WriteExcelNameToPath(string excelFloder,Action<string, float> readCallback = null)
        {
            Dictionary<string, string> nameToPath = new Dictionary<string, string>();

            mExcelSheetCaches = new Dictionary<string, ExcelWorksheet>();
            if (Directory.Exists(excelFloder))
            {
                var dirInfo = Directory.CreateDirectory(excelFloder);

                FileInfo[] fileInfos = dirInfo.GetFiles();
                for (int i = 0; i < fileInfos.Length; i++)
                {
                    FileInfo fileInfo = fileInfos[i];
                    if (fileInfo.Name.EndsWith(".xlsx"))
                    {
                        try
                        {
                            ExcelPackage ep = new ExcelPackage(fileInfo);
                            ExcelWorksheets worksheets = ep.Workbook.Worksheets;
                            for (int j = 1; j <= worksheets.Count; j++)
                            {
                                ExcelWorksheet sheet = worksheets[j];
                                if (nameToPath.ContainsKey(sheet.Name))
                                {
                                    string oldExcel = nameToPath[sheet.Name];
                                    UnityEngine.Debug.LogErrorFormat(" 配置文件【{0}】和【{1}】都存在表 {2}！", oldExcel, fileInfo.Name, sheet.Name);
                                }
                                else
                                {
                                    nameToPath.Add(sheet.Name, fileInfo.Name);
                                }
                            }
                        }
                        catch (System.Exception ex)
                        {
                            UnityEngine.Debug.LogErrorFormat(" 读取配置表:{0}报错！,MSG = {1}!", fileInfo.Name, ex.Message);
                        }
                    }
                    readCallback?.Invoke(fileInfo.Name, (i + 1) * 1f / fileInfos.Length);
                }
            }
            var json = UnityEngine.JsonUtility.ToJson(nameToPath);
            File.WriteAllText(Path.Combine(excelFloder, NAME_TO_PATH), json.ToString());
        }

        public static Dictionary<string, ExcelWorksheet> ReadALLExcelSheets(string excelFloder,
            Action<string, float> readCallback = null)
        {
            mExcelSheetCaches = new Dictionary<string, ExcelWorksheet>();
            if (Directory.Exists(excelFloder))
            {
                var dirInfo = Directory.CreateDirectory(excelFloder);

                FileInfo[] fileInfos = dirInfo.GetFiles();
                for (int i = 0; i < fileInfos.Length; i++)
                {
                    FileInfo fileInfo = fileInfos[i];
                    if (fileInfo.Name.EndsWith(".xlsx"))
                    {
                        var sheelLst = ReadExcelSheets(fileInfo);
                        if (sheelLst == null) continue;
                        for (int j = 0; j < sheelLst.Count; j++)
                        {
                            ExcelWorksheet sheet = sheelLst[j];
                            int sheetRowNum = 0, sheetColuNum = 0;
                            if (sheet.Dimension != null)
                            {
                                sheetRowNum = sheet.Dimension.Rows;
                                sheetColuNum = sheet.Dimension.Columns;
                            }
                            if (sheetRowNum >= TranslatorTable.START_ROW + 3)
                            {
                                mExcelSheetCaches.Add(sheet.Name, sheet);
                            }
                            else
                            {
                                ;// UnityEngine.Debug.LogErrorFormat("存在无效配置表:[{0}-{1}],请确认！", fileInfo.Name, sheet.Name);
                            }
                        }
                    }
                    readCallback?.Invoke(fileInfo.Name, (i+1)*1f/fileInfos.Length);
                }
            }
            return mExcelSheetCaches;
        }

        public static bool WriteExcelSheets(string filePath, List<ExcelWorksheet> sheelLst)
        {
            if (File.Exists(filePath) && filePath.EndsWith(".xlsx"))
            {
                try
                {
                    FileInfo output = new FileInfo(filePath);
                    ExcelPackage ep = new ExcelPackage(output);
                    //for (int i = 0; i < sheelLst.Count; i++)
                    //{
                    //    ExcelWorksheet sheet = ep.Workbook.Worksheets.Add(table.TableName);
                    //    for (int row = 1; row <= table.NumberOfRows; row++)
                    //    {
                    //        for (int column = 1; column <= table.NumberOfColumns; column++)
                    //        {
                    //            sheet.Cells[row, column].Value = table.GetValue(row, column);
                    //        }
                    //    }
                    //}
                    ep.SaveAs(output);
                    return true;
                }
                catch (System.Exception ex)
                {
                    throw new System.Exception(ex.Message);
                }
            }
            return false;
        }

        public static ExcelWorksheet ReadExcelSheet(string excelFloder, string sheetName)
        {
            if (mExcelSheetCaches == null) mExcelSheetCaches = new Dictionary<string, ExcelWorksheet>();

            ExcelWorksheet sheet;
            if (!mExcelSheetCaches.TryGetValue(sheetName, out sheet))
            {
                var json = File.ReadAllText(Path.Combine(excelFloder, NAME_TO_PATH));
                var nameToPath = UnityEngine.JsonUtility.FromJson<Dictionary<string, string>>(json);
                string excelName = string.Empty;
                if (nameToPath.TryGetValue(sheetName, out excelName))
                {
                    var dirInfo = Directory.CreateDirectory(excelFloder);
                    FileInfo[] fileInfos = dirInfo.GetFiles();
                    for (int i = 0; i < fileInfos.Length; i++)
                    {
                        FileInfo fileInfo = fileInfos[i];
                        if (fileInfo.Name == excelName)
                        {
                            ExcelPackage ep = new ExcelPackage(fileInfo);
                            ExcelWorksheets worksheets = ep.Workbook.Worksheets;
                            for (int j = 1; j <= worksheets.Count; j++)
                            {
                                ExcelWorksheet newSheet = worksheets[j];
                                mExcelSheetCaches.Add(newSheet.Name, newSheet);
                                if (newSheet.Name == sheetName) sheet = newSheet;
                            }
                        }
                    }
                }
                else
                {
                    UnityEngine.Debug.LogErrorFormat("新增了配置表【{0}】， 请用菜单工具生成一下映射。", sheetName);
                }
            }
            return sheet;
        }

        public static void GenerateBytesByExcel(string excelFloder, string bytesFloder, int readMask = 0x7FFFFFFF)
        {
            if (!Directory.Exists(bytesFloder)) Directory.CreateDirectory(bytesFloder);

            ReadALLExcelSheets(excelFloder);
            foreach (var excelSheet in mExcelSheetCaches.Values)
            {
                var translator = new TranslatorTable(excelSheet, readMask);
                string byte_path = Path.Combine(bytesFloder, translator.sheetName + ".byte");
                File.WriteAllBytes(byte_path, translator.ToDataEntryBytes());
            }
        }

        public static void GenerateJsonByExcel(string excelFloder, string jsonFloder, int readMask = 0x7FFFFFFF)
        {
            if (!Directory.Exists(jsonFloder)) Directory.CreateDirectory(jsonFloder);

            ReadALLExcelSheets(excelFloder);
            foreach (var excelSheet in mExcelSheetCaches.Values)
            {
                var translator = new TranslatorTable(excelSheet, readMask);
                string json_path = Path.Combine(jsonFloder, translator.sheetName + ".json");
                File.WriteAllBytes(json_path, UTF8Encoding.UTF8.GetBytes(translator.ToJson()));
            }
        }

        public static void GenerateEntryClassByExcel(string excelFloder, string codeFloder, int readMask = 0x7FFFFFFF)
        {
            if (!Directory.Exists(codeFloder)) Directory.CreateDirectory(codeFloder);
            StringBuilder stringBuilder = new StringBuilder();

            ReadALLExcelSheets(excelFloder);
            foreach (var excelSheet in mExcelSheetCaches.Values)
            {
                stringBuilder.Clear();
                var translator = new TranslatorTable(excelSheet, readMask);
                string path = Path.Combine(codeFloder, translator.sheetName + ".cs");
                File.WriteAllBytes(path, UTF8Encoding.UTF8.GetBytes(translator.ToDataEntryClass()));
            }
        }

        public static void GenerateLuaTableByExcel(string excelFloder, string luaFloder, int readMask = 0x7FFFFFFF)
        {
            if (!Directory.Exists(luaFloder)) Directory.CreateDirectory(luaFloder);

            ReadALLExcelSheets(excelFloder);
            foreach (var excelSheet in mExcelSheetCaches.Values)
            {
                var translator = new TranslatorTable(excelSheet, readMask);
                string path = Path.Combine(luaFloder, translator.sheetName + ".lua");
                File.WriteAllBytes(path, UTF8Encoding.UTF8.GetBytes(translator.ToLuaTable()));
            }
        }

        public static string SheetNameToDataEntryClassName(string sheetName)
        {
            return string.Format("DataEntry_{0}", sheetName);
        }

        public static string DataEntryClassNameToSheetName(string className)
        {
            return className.Substring("DataEntry_".Length);
        }

        public static string GetDataEntryClassFullName(string sheetName)
        {
            return string.Format("data.DataEntry_{0}", sheetName);
        }

        #endregion

        #region Encrypt Methods

        //临时的8位数密码
        public static byte[] XOR = new byte[] { 0x11, 0x01, 0xFA, 0xBC, 0x57, 0xFF, 0xF1, 0xEA };

        public static void EncryptBytes(byte[] xor, ref byte[] __bytes)
        {
            for (int i = 0; i < __bytes.Length; i++)
            {
                byte xor_byte = xor[i % xor.Length];
                __bytes[i] = (byte)(__bytes[i] ^ xor_byte);
            }
        }

        #endregion
    }
}
