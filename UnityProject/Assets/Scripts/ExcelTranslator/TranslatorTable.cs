using System.Collections.Generic;
using System.Collections;
using System;
using System.Text;
using OfficeOpenXml;
using System.Reflection;

namespace Engine.Core.ExcelTranslator
{
    /// <summary>
    /// 中转数据表
    /// </summary>
    public sealed class TranslatorTable
    {
        public const int START_ROW = 3;

        public readonly string sheetName;

        public readonly int nRow = 0;

        public readonly int nColu = 0;

        private string[] mIds = null;

        private string[] mAttriNames = null;

        private ValueType[] mAttriTypes = null;

        private List<byte[]>[,] mCellValues = null;

        /// <summary>
        /// Excel表转换为Table. 1-2两行没用
        /// 注：第4行是属性: string.
        ///     第3行是读取方式: readFlag读取标记 按位存取标记 一共 0 - 31 种方式.
        ///     第5行是数据类型: 1=int\ 2=float\ 3=bool\ 0=string.
        ///     第6行开始是数据: string.
        ///     第1列必须为id: string.
        /// </summary>
        public TranslatorTable(ExcelWorksheet excelSheet, int readMask)
        {
            try
            {
                //解析数据表有效数据行列
                sheetName = excelSheet.Name;
                int sheetRowNum = 0, sheetColuNum = 0;
                if (excelSheet.Dimension != null)
                {
                    sheetRowNum = excelSheet.Dimension.Rows;
                    sheetColuNum = excelSheet.Dimension.Columns;
                }

                //第3-5行是读取方式 Int（确定列数）、值类型、值属性名称
                List<int> _needReadColus = new List<int>();
                List<ValueType> _ValueTypes = new List<ValueType>();
                List<string> _AttriNames = new List<string>();

                for (int colu = 1; colu <= sheetColuNum; colu++)
                {
                    var readTagObj = excelSheet.Cells[START_ROW, colu].Value;
                    var valueTypeObj = excelSheet.Cells[START_ROW + 1, colu].Value;
                    var attriNameObj = excelSheet.Cells[START_ROW + 2, colu].Value;
                    var readMaskString = readTagObj != null ? readTagObj.ToString() : string.Empty;
                    var valueTypeString = valueTypeObj != null ? valueTypeObj.ToString() : string.Empty;
                    var attriNameString = attriNameObj != null ? attriNameObj.ToString() : string.Empty;
                    if (!string.IsNullOrEmpty(readMaskString) && 
                        !string.IsNullOrEmpty(valueTypeString) && 
                        !string.IsNullOrEmpty(attriNameString))
                    {
                        int _ReadMask = 3;
                        int.TryParse(readMaskString, out _ReadMask);
                        if (_ReadMask != 3) //|| (readMask & _ReadMask) > 0)
                        {
                            int indexOf = attriNameString.IndexOf("#[]");
                            if (indexOf != -1)
                                attriNameString = attriNameString.Substring(0, indexOf);
                            _AttriNames.Add(attriNameString);
                            if (colu == 1) valueTypeString = "0";
                            _ValueTypes.Add(ExcelValueToValueType(valueTypeString, indexOf != -1));
                            _needReadColus.Add(colu);
                        }
                    }
                    else
                    {
                        //UnityEngine.Debug.LogWarningFormat("解析Excel：[{0}]表，第{1}列为空! 忽略了该列。", sheetName, colu);
                    }
                }
                nColu = _needReadColus.Count;
                mAttriNames = _AttriNames.ToArray();
                mAttriTypes = _ValueTypes.ToArray();

                //检测哪些行数据是有效可以读取的
                List<int> _neetReadRows = new List<int>();
                for (int rowNum = START_ROW + 3; rowNum <= sheetRowNum; rowNum++)
                {
                    var idValue = excelSheet.Cells[rowNum, 1].Value;
                    string id = idValue != null ? idValue.ToString() : string.Empty;
                    if (!string.IsNullOrEmpty(id))
                    {
                        _neetReadRows.Add(rowNum);
                    }
                    else
                    {
                        //UnityEngine.Debug.LogWarningFormat("解析Excel：[{0}]表，第{1}行为空! 忽略了该行。", sheetName, rowNum);
                    }
                }

                //从第6行开始才是数据内容
                nRow = _neetReadRows.Count;
                mIds = new string[nRow];
                mCellValues = new List<byte[]>[nRow, nColu];
                for (int index_r = 0; index_r < nRow; index_r++)
                {
                    int rowNum = _neetReadRows[index_r];
                    for (int index_c = 0; index_c < nColu; index_c++)
                    {
                        int coluNum = _needReadColus[index_c];
                        string value = string.Empty;
                        if (excelSheet.Cells[rowNum, coluNum].Value != null)
                        {
                            value = excelSheet.Cells[rowNum, coluNum].Value.ToString();
                        }
                        mCellValues[index_r, index_c] = StringToByteList(mAttriTypes[index_c], value);
                    }
                    mIds[index_r] = UTF8Encoding.UTF8.GetString(mCellValues[index_r, 0][0]);
                }
            }
            catch (Exception ex)
            {
                UnityEngine.Debug.LogErrorFormat("解析Excel：[{0}]表错误! Msg = {1}!", sheetName, ex.Message);
            }
        }

        /// <summary>
        /// bytes字节 转换为Table
        /// </summary>
        public TranslatorTable(byte[] bytes)
        {
            var __temp_buff = new ExcelTranslatorBuffer(bytes, (UInt32)bytes.Length);

            //写读出格的行列数
            __temp_buff.Out(out sheetName);
            __temp_buff.Out(out nRow);
            __temp_buff.Out(out nColu);

            //读取属性字段 和属性数据类型
            mAttriNames = new string[nColu];
            mAttriTypes = new ValueType[nColu];
            for (int index_c = 0; index_c < nColu; index_c++)
            {
                string attr_name = string.Empty;
                __temp_buff.Out(out attr_name);
                mAttriNames[index_c] = attr_name;

                int attr_type = 0;
                __temp_buff.Out(out attr_type);
                mAttriTypes[index_c] = (ValueType)attr_type;
            }

            //读取数据
            mIds = new string[nRow];
            mCellValues = new List<byte[]>[nRow, nColu];
            for (int index_r = 0; index_r < nRow; index_r++)
            {
                for (int index_c = 0; index_c < nColu; index_c++)
                {
                    List<byte[]> value;
                    __temp_buff.Out(mAttriTypes[index_c], out value);
                    mCellValues[index_r, index_c] = value;
                }
                mIds[index_r] = UTF8Encoding.UTF8.GetString(mCellValues[index_r, 0][0]);
            }
        }

        public string ID(int index_r)
        {
            index_r = assert_row(index_r);
            return mIds[index_r];
        }

        public string Attr(int index_c)
        {
            index_c = assert_colu(index_c);
            return mAttriNames[index_c];
        }

        public ValueType Type(int index_c)
        {
            index_c = assert_colu(index_c);
            return mAttriTypes[index_c];
        }

        public List<byte[]> Value(int index_r, int index_c)
        {
            return mCellValues[assert_row(index_r), assert_colu(index_c)];
        }

        private int assert_row(int index_r)
        {
            if (index_r < 0 || index_r >= nRow)
            {
                index_r = Math.Max(0, Math.Min(index_r, nRow - 1));
                throw new Exception("assert_row 索引越界!");
            }
            return index_r;
        }

        private int assert_colu(int index_c)
        {
            if (index_c < 0 || index_c >= nColu)
            {
                index_c = Math.Max(0, Math.Min(index_c, nColu - 1));
                throw new Exception("assert_colu 索引越界!");
            }
            return index_c;
        }

        #region 文件转换接口

        public string ToJson()
        {
            StringBuilder stringBuilder = new StringBuilder();
            stringBuilder.Append("{\n");
            for (int index_r = 0; index_r < nRow; index_r++)
            {
                string id = ID(index_r);
                stringBuilder.Append("  \"");
                stringBuilder.Append(id);
                stringBuilder.Append("\":{");
                for (int index_c = 0; index_c < nColu; index_c++)
                {
                    stringBuilder.Append("\"");
                    stringBuilder.Append(Attr(index_c));
                    stringBuilder.Append("\":");

                    var valueType = mAttriTypes[index_c];
                    List<byte[]> cellValue = mCellValues[index_r, index_c];

                    //非数组
                    if (valueType == ValueType.Int32 || valueType == ValueType.Bool || 
                        valueType == ValueType.Float || valueType == ValueType.String)
                    {
                        stringBuilder.Append("\"");
                        string __string = string.Empty;
                        switch (valueType)
                        {
                            case ValueType.Int32:
                                __string = BitConverter.ToInt32(cellValue[0], 0).ToString();
                                break;
                            case ValueType.Bool:
                                __string = BitConverter.ToBoolean(cellValue[0], 0).ToString();
                                break;
                            case ValueType.Float:
                                __string = BitConverter.ToSingle(cellValue[0], 0).ToString();
                                break;
                            case ValueType.String:
                                __string = UTF8Encoding.UTF8.GetString(cellValue[0]);
                                __string = __string.Replace("\"", "\\\"");
                                break;
                        }
                        stringBuilder.Append(__string);
                        stringBuilder.Append("\"");
                    }
                    else
                    {
                        stringBuilder.Append("[");
                        int length = cellValue.Count;
                        for (int i = 0; i < length; i++)
                        {
                            stringBuilder.Append("\"");
                            string __string = string.Empty;
                            switch (valueType)
                            {
                                case ValueType.Int32Array:
                                    __string = BitConverter.ToInt32(cellValue[i], 0).ToString();
                                    break;
                                case ValueType.BoolArray:
                                    __string = BitConverter.ToBoolean(cellValue[i], 0).ToString();
                                    break;
                                case ValueType.FloatArray:
                                    __string = BitConverter.ToSingle(cellValue[i], 0).ToString();
                                    break;
                                case ValueType.StringArray:
                                    __string = UTF8Encoding.UTF8.GetString(cellValue[i]);
                                    __string = __string.Replace("\"", "\\\"");
                                    break;
                            }
                            stringBuilder.Append(__string);
                            stringBuilder.Append("\"");
                            if (i < length - 1) stringBuilder.Append(",");
                        }
                        stringBuilder.Append("]");
                    }
                    if (index_c < nColu - 1) stringBuilder.Append(",");
                }
                stringBuilder.Append("}");
                if (index_r < nRow - 1) stringBuilder.Append(",");
                stringBuilder.Append("\n");
            }
            stringBuilder.Append("}");
            return stringBuilder.ToString();
        }

        public byte[] ToDataEntryBytes()
        {
            var __buffer = new ExcelTranslatorBuffer();
            __buffer.Reset();

            //写入表格的行列数
            __buffer.In(sheetName);
            __buffer.In(nRow);
            __buffer.In(nColu);

            //属性写入字节流
            for (int i = 0; i < nColu; i++)
            {
                __buffer.In(mAttriNames[i]);
                __buffer.In((int)mAttriTypes[i]);
            }
            //遍历所有的数据行，写入字节流
            for (int i = 0; i < nRow; i++)
            {
                for (int j = 0; j < nColu; j++)
                {
                    __buffer.In(mAttriTypes[j], mCellValues[i, j]);
                }
            }
            byte[] __bytes = new byte[__buffer.Size];
            Array.Copy(__buffer.GetBuffer(), __bytes, __buffer.Size);
            return __bytes;
        }

        public string ToDataEntryClass()
        {
            StringBuilder stringBuilder = new StringBuilder();

            stringBuilder.Append("//File Generate By ExcelTranslator, Don't Modify It!\n");
            stringBuilder.Append("using System;\n");
            stringBuilder.Append("using Engine.Core.ExcelTranslator;\n\n");
            stringBuilder.Append("namespace data\n");
            stringBuilder.Append("{\n");
            stringBuilder.Append(string.Format("\tpublic class {0} : DataEntryBase\n",
                ExcelTranslatorUtility.SheetNameToDataEntryClassName(sheetName)));
            stringBuilder.Append("\t{\n");

            stringBuilder.Append(string.Format("\t\tpublic static string sheetName = \"{0}\";\n", sheetName));

            //字段
            for (int i = 0; i < nColu; i++)
            {
                string attr_name = Attr(i);
                string type_name = GetCSharpTypeName(Type(i));
                stringBuilder.Append(string.Format("\t\tpublic {0} {1};\n", type_name, attr_name));
            }

            //bytes反序列化函数
            stringBuilder.Append("\n\t\tpublic override void DeSerialized(ExcelTranslatorBuffer buffer)\n");
            stringBuilder.Append("\t\t{\n");
            for (int i = 0; i < nColu; i++)
            {
                string attr_name = Attr(i);
                stringBuilder.Append(string.Format("\t\t\tbuffer.Out(out {0});\n", attr_name));
            }
            stringBuilder.Append(string.Format("\t\t\tKEY = {0}.ToString();\n", Attr(0)));
            stringBuilder.Append("\t\t}\n");

            stringBuilder.Append("\t}\n");
            stringBuilder.Append("}");
            return stringBuilder.ToString();
        }

        private string GetCSharpTypeName(ValueType valueType)
        {
            switch (valueType)
            {
                case ValueType.Int32:
                    return "int";
                case ValueType.Bool:
                    return "bool";
                case ValueType.Float:
                    return "float";
                case ValueType.String:
                    return "string";
                case ValueType.Int32Array:
                    return "int[]";
                case ValueType.BoolArray:
                    return "bool[]";
                case ValueType.FloatArray:
                    return "float[]";
                case ValueType.StringArray:
                    return "string[]";
                default:
                    throw new Exception("ExcelTranslatorBuffer.OutDynamicValue() 不存在的类型！ " + valueType);
            }
        }

        public string ToLuaTable()
        {
            StringBuilder stringBuilder = new StringBuilder();
            stringBuilder.Append("-- File Generate By ExcelTranslator\n");
            stringBuilder.Append(string.Format("local {0} = \n", sheetName));
            stringBuilder.Append("{\n");

            for (int index_r = 0; index_r < nRow; index_r++)
            {
                stringBuilder.Append(string.Format("  [\"{0}\"] = ", ID(index_r)));
                stringBuilder.Append("{");

                for (int index_c = 0; index_c < nColu; index_c++)
                {
                    string attrName = Attr(index_c);
                    var valueType = mAttriTypes[index_c];
                    List<byte[]> cellValue = mCellValues[index_r, index_c];

                    //非数组
                    if (valueType == ValueType.Int32 || valueType == ValueType.Bool ||
                        valueType == ValueType.Float || valueType == ValueType.String)
                    {
                        string __string = string.Empty;
                        switch (valueType)
                        {
                            case ValueType.Int32:
                                __string = BitConverter.ToInt32(cellValue[0], 0).ToString();
                                break;
                            case ValueType.Bool:
                                var __bool = BitConverter.ToBoolean(cellValue[0], 0);
                                __string = __bool ? "true" : "false";
                                break;
                            case ValueType.Float:
                                __string = BitConverter.ToSingle(cellValue[0], 0).ToString();
                                break;
                            case ValueType.String:
                                __string = UTF8Encoding.UTF8.GetString(cellValue[0]);
                                __string = __string.Replace("\"", "\\\"");
                                __string = string.Format("\"{0}\"", __string);
                                break;
                        }
                        stringBuilder.Append(string.Format("{0}={1}", attrName, __string));
                    }
                    else
                    {
                        stringBuilder.Append(attrName + "={");
                        int length = cellValue.Count;
                        for (int i = 0; i < length; i++)
                        {
                            string __string = string.Empty;
                            switch (valueType)
                            {
                                case ValueType.Int32Array:
                                    __string = BitConverter.ToInt32(cellValue[i], 0).ToString();
                                    break;
                                case ValueType.BoolArray:
                                    var __bool = BitConverter.ToBoolean(cellValue[i], 0);
                                    __string = __bool ? "true" : "false";
                                    break;
                                case ValueType.FloatArray:
                                    __string = BitConverter.ToSingle(cellValue[i], 0).ToString();
                                    break;
                                case ValueType.StringArray:
                                    __string = UTF8Encoding.UTF8.GetString(cellValue[i]);
                                    __string = __string.Replace("\"", "\\\"");
                                    __string = string.Format("\"{0}\"", __string);
                                    break;
                            }
                            stringBuilder.Append(__string);
                            if (i < length - 1) stringBuilder.Append(",");
                        }
                        stringBuilder.Append("}");
                    }
                    if (index_c < nColu - 1) stringBuilder.Append(",");
                }
                stringBuilder.Append("}");
                if (index_r < nRow - 1) stringBuilder.Append(",");
                stringBuilder.Append("\n");
            }
            stringBuilder.Append("}\n");
            stringBuilder.Append(string.Format("return {0};", sheetName));
            return stringBuilder.ToString();
        }

        #endregion

        #region 数据表 静态方法

        private static ExcelTranslatorBuffer __temp_buff = new ExcelTranslatorBuffer();

        public static DataEntryCache ToTableCache(byte[] bytes, Type type)
        {
            __temp_buff.Reset();
            __temp_buff.Append(bytes, (uint)bytes.Length);

            //1 读出格的行列数
            string sheetName; int nRow, nColu;
            __temp_buff.Out(out sheetName).Out(out nRow).Out(out nColu);

            //new一个实例
            var tableCache = new DataEntryCache(sheetName, nRow, nColu);

            //2 读出属性字段 和属性数据类型
            for (int index_c = 0; index_c < nColu; index_c++)
            {
                string attriname = string.Empty;
                int attritype = 0;
                __temp_buff.Out(out attriname).Out(out attritype);
            }

            //反序列化表对象
            for (int index_r = 0; index_r < nRow; index_r++)
            {
                var entry = Activator.CreateInstance(type) as DataEntryBase;
                entry.DeSerialized(__temp_buff);
                tableCache[entry.KEY] = entry;
            }
            return tableCache;
        }

        public static DataEntryCache ToTableCache(ExcelWorksheet excelSheet, int readMask, Type type)
        {
            TranslatorTable table = new TranslatorTable(excelSheet, readMask);
            return ToTableCache(table.ToDataEntryBytes(), type);
        }

        public static string ToJson(byte[] bytes)
        {
            TranslatorTable table = new TranslatorTable(bytes);
            return table.ToJson();
        }

        public static string ToJson(ExcelWorksheet excelSheet, int readMask)
        {
            TranslatorTable table = new TranslatorTable(excelSheet, readMask);
            return table.ToJson();
        }

        public static string ToLuaLable(byte[] bytes)
        {
            TranslatorTable table = new TranslatorTable(bytes);
            return table.ToLuaTable();
        }

        public static string ToLuaLable(ExcelWorksheet excelSheet, int readMask)
        {
            TranslatorTable table = new TranslatorTable(excelSheet, readMask);
            return table.ToLuaTable();
        }

        public static ValueType StringToValueType(string typeString)
        {
            typeString = typeString.ToLower();
            if (typeString == "int")
                return ValueType.Int32;
            else if (typeString == "bool")
                return ValueType.Bool;
            else if (typeString == "float")
                return ValueType.Float;
            else if (typeString == "string")
                return ValueType.String;
            else if (typeString == "int[]")
                return ValueType.Int32Array;
            else if (typeString == "bool[]")
                return ValueType.BoolArray;
            else if (typeString == "float[]")
                return ValueType.FloatArray;
            else if (typeString == "string[]")
                return ValueType.StringArray;
            else
                throw new Exception("StringToValueType(): 属性的值类型错误!" + typeString);
        }

        public static List<byte[]> StringToByteList(ValueType type, string value)
        {
            try
            {
                if (type == ValueType.Int32)
                {
                    int __vInt32 = 0;
                    int.TryParse(value, out __vInt32);
                    return new List<byte[]>{BitConverter.GetBytes(__vInt32)};
                }
                else if (type == ValueType.Float)
                {
                    float __vFloat = 0;
                    float.TryParse(value, out __vFloat);
                    return new List<byte[]> { BitConverter.GetBytes(__vFloat) };
                }
                else if (type == ValueType.Bool)
                {
                    bool __vBool = true;
                    bool.TryParse(value, out __vBool);
                    return new List<byte[]> { BitConverter.GetBytes(__vBool) };
                }
                else if (type == ValueType.String)
                {
                    return new List<byte[]> { Encoding.UTF8.GetBytes(value) };
                }
                else
                {
                    var valueArray = string.IsNullOrEmpty(value) ? new string[0] : value.Split('|');
                    if (type == ValueType.Int32Array)
                    {
                        var resultArray = new List<byte[]>();
                        for (int i = 0; i < valueArray.Length; i++)
                        {
                            int __vInt32 = 0;
                            int.TryParse(valueArray[i], out __vInt32);
                            resultArray.Add(BitConverter.GetBytes(__vInt32));
                        }
                        return resultArray;
                    }
                    else if (type == ValueType.FloatArray)
                    {
                        var resultArray = new List<byte[]>();
                        for (int i = 0; i < valueArray.Length; i++)
                        {
                            float __vFloat = 0;
                            float.TryParse(valueArray[i], out __vFloat);
                            resultArray.Add(BitConverter.GetBytes(__vFloat));
                        }
                        return resultArray;
                    }
                    else if (type == ValueType.BoolArray)
                    {
                        var resultArray = new List<byte[]>();
                        for (int i = 0; i < valueArray.Length; i++)
                        {
                            bool __vBool = true;
                            bool.TryParse(valueArray[i], out __vBool);
                            resultArray.Add(BitConverter.GetBytes(__vBool));
                        }
                        return resultArray;
                    }
                    else if (type == ValueType.StringArray)
                    {
                        var resultArray = new List<byte[]>();
                        for (int i = 0; i < valueArray.Length; i++)
                        {
                            resultArray.Add(Encoding.UTF8.GetBytes(valueArray[i]));
                        }
                        return resultArray;
                    }
                }
                throw new Exception(string.Format("属性类型[{0}]错误！", type));
            }
            catch (Exception e)
            {
                throw new Exception(string.Format("属性类型[{0}]错误！Error Msg = {1}", type, e.Message));
            }
        }

        public static ValueType ExcelValueToValueType(string type, bool isArray)
        {
            if (type == "1")
            {
                return isArray ? ValueType.Int32Array : ValueType.Int32;
            }
            else if (type == "2")
            {
                return isArray ? ValueType.FloatArray : ValueType.Float;
            }
            else if (type == "3")
            {
                return isArray ? ValueType.BoolArray : ValueType.Bool;
            }
            return isArray ? ValueType.StringArray : ValueType.String;
        }

        #endregion
    }
}
