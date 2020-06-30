using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace Engine.Core.ExcelTranslator
{
    [Serializable]
    public abstract class DataEntryBase
    {
        public string KEY;
        public abstract void DeSerialized(ExcelTranslatorBuffer buffer);
    }
}
