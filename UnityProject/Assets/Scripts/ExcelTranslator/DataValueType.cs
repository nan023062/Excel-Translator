using System.Collections.Generic;
using System.Collections;
using System;
using System.Text;
using OfficeOpenXml;
using System.Reflection;

namespace Engine.Core.ExcelTranslator
{
    public enum ValueType
    {
        Int32,
        Bool,
        Float,
        String,
        Int32Array,
        BoolArray,
        FloatArray,
        StringArray,
    }
}
