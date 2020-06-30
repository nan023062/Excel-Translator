using System;
using System.Collections.Generic;
using System.Text;
using UnityEngine;
using System.Reflection;
using OfficeOpenXml;
using Engine.Core.ExcelTranslator;

namespace Sample
{
    public class ConfigSample
    {
        private Dictionary<string, DataEntryCache> mDataEntryCaches = null;

        private Dictionary<string, ExcelWorksheet> mCachedExcels = null;

        public void Init()
        {
            mDataEntryCaches = new Dictionary<string, DataEntryCache>();
            mCachedExcels = new Dictionary<string, ExcelWorksheet>();
        }

        private byte[] LoadConfigBytes(string configName)
        {
            byte[] bytes = null;



            return bytes;
        }

        public T GetConfig<T>(string id) where T : DataEntryBase
        {
            Type type = typeof(T);
            var fi = type.GetField("sheetName", BindingFlags.Static| BindingFlags.Public);
            string configName = fi.GetValue(type).ToString();
            return GetTableCache(configName, type).GetEntry<T>(id);
        }

        public DataEntryCache GetTableCache<T>() where T : DataEntryBase
        {
            Type type = typeof(T);
            var fi = type.GetField("sheetName", BindingFlags.Static | BindingFlags.Public);
            string configName = fi.GetValue(type).ToString();
            return GetTableCache(configName, type);
        }

        public DataEntryCache GetTableCache(string configName, Type type)
        {
            DataEntryCache entryCache = null;
            if (!mDataEntryCaches.TryGetValue(configName, out entryCache))
            {
                byte[] bytes = LoadConfigBytes(configName);
                entryCache = TranslatorTable.ToTableCache(bytes, type);
                mDataEntryCaches.Add(configName, entryCache);
            }
            return entryCache;
        }

        public byte[] GetLuaTableBytes(string configName)
        {
            byte[] bytes = LoadConfigBytes(configName);
            string luaString = TranslatorTable.ToLuaLable(bytes);
            return UTF8Encoding.UTF8.GetBytes(luaString);
        }

        public string GetJsonDataTable(string configName)
        {
            byte[] bytes = LoadConfigBytes(configName);
            return TranslatorTable.ToJson(bytes);
        }

        public ExcelWorksheet ReadExcelSheet(string configName)
        {
            ExcelWorksheet sheet;
            if (!mCachedExcels.TryGetValue(configName, out sheet))
            {
                string excelFloder = string.Empty;
                sheet = ExcelTranslatorUtility.ReadExcelSheet(excelFloder, configName);
                mCachedExcels.Add(configName, sheet);
            }
            return sheet;
        }
    }

}