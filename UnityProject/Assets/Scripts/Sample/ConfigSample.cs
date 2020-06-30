using System;
using System.Collections.Generic;
using System.Text;
using System.Reflection;
using Engine.Core.ExcelTranslator;
using System.IO;
using UnityEngine;
using data;

namespace Sample
{
    public class ConfigSample : MonoBehaviour
    {
        private Dictionary<string, DataEntryCache> mDataEntryCaches = null;

        public void Awake()
        {
            mDataEntryCaches = new Dictionary<string, DataEntryCache>();
        }

        public DataEntry_sound_cfg sound;

        public DataEntry_EffectConfig effect;

        public void Start()
        {
            var soundCfg = GetTableCache<DataEntry_sound_cfg>();

            sound = GetConfig<DataEntry_sound_cfg>("100001");

            var effectCfg = GetTableCache<DataEntry_EffectConfig>();

            effect = GetConfig<DataEntry_EffectConfig>("400023");
        }

        private byte[] LoadConfigBytes(string configName)
        {
            string path = Path.Combine(Engine.Config.DataEditorTool.ByteFilePath, configName + ".byte");
            return File.ReadAllBytes(path);
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

    }

}