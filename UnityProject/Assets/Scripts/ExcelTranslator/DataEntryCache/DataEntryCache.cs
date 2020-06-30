using System;
using System.Collections.Generic;

namespace Engine.Core.ExcelTranslator
{
    public class DataEntryCache
    {
        public readonly string sheetName;
        public readonly int nRow;
        public readonly int nColu;

        private Dictionary<string, DataEntryBase> mTableEntries;

        public Dictionary<string, DataEntryBase> TableEntries { get { return mTableEntries; } }

        public int Count { get { return mTableEntries.Count; } }

        public DataEntryBase this[string id] 
        {
            get
            {
                DataEntryBase result = null;
                mTableEntries.TryGetValue(id, out result);
                return result;
            }
            set
            {
                mTableEntries[id] = value;
            }
        }

        public List<string> GetEntryIDList()
        {
            return new List<string>(mTableEntries.Keys);
        }

        public T GetEntry<T>(string id) where T : DataEntryBase
        {
            return this[id] as T;
        }

        public DataEntryCache(string sheetName, int nRow, int nColu)
        {
            this.sheetName = sheetName;
            this.nRow = nRow;
            this.nColu = nColu;
            mTableEntries = new Dictionary<string, DataEntryBase>();
        }
    }
}
