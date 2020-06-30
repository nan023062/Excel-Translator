using System;
using System.Collections.Generic;
using System.Text;

namespace Engine.Core.ExcelTranslator
{
    public class ExcelTranslatorBuffer
    {
        public const UInt32 MAXSIZE = 0xffffffff;
        protected const int DEFAULT_BUFF_SIZE = 256 << 3;
        private byte[] __InOut_buf = new byte[128];
        private UInt32 m_rpos = 0;
        private UInt32 m_wpos = 0;
        private UInt32 m_validateSize = 0;
        protected UInt32 m_buffSize = 0;
        protected Byte[] m_buff = new Byte[DEFAULT_BUFF_SIZE];

        //属性
        public UInt32 Size { get { return m_validateSize; } set { m_validateSize = value; } }
        public UInt32 ReadPosition { get { return m_rpos; } }
        public UInt32 WritePosition { get { return m_wpos; } }

        //构造
        public ExcelTranslatorBuffer()
        {
            m_validateSize = m_rpos = m_wpos = 0;
            _Resize(DEFAULT_BUFF_SIZE);
        }

        public ExcelTranslatorBuffer(ExcelTranslatorBuffer r)
        {
            m_validateSize = m_rpos = m_wpos = 0;
            Resize(r.m_validateSize);
        }

        public ExcelTranslatorBuffer(Byte[] buf, UInt32 size)
        {
            m_buffSize = 0;
            m_validateSize = m_rpos = m_wpos = 0;
            if (buf != null)
            {
                Resize(size);
                Array.Copy(buf, m_buff, size);
            }
            else
            {
                _Resize(size);
            }
        }

        public void Reset()
        {
            m_buffSize = m_rpos = m_wpos = 0;
        }

        private void _Resize(UInt32 newsize)
        {
            if (newsize > m_buffSize)
            {
                newsize = (newsize > m_buffSize * 2) ? newsize : (m_buffSize * 2);
                if ((newsize > m_buff.Length * sizeof(Byte)))// && validateSize > 0)
                {
                    if (m_validateSize > 0)
                    {
                        Byte[] tempbuff = m_buff; // copy the content to the tempbuff
                        m_buff = new Byte[newsize];
                        Array.Copy(tempbuff, m_buff, m_validateSize);
                        tempbuff = null;
                    }
                    else
                    {
                        m_buff = new Byte[newsize];
                    }
                }
                m_buffSize = newsize;
            }
        }

        public void Resize(UInt32 newsize)
        {
            _Resize(newsize);
            if (newsize > m_validateSize)
                m_validateSize = newsize;
        }

        public byte[] GetBuffer() { return m_buff; }

        public UInt32 GetBuffSize() { return m_buffSize; }

        //读写接口
        void _Write(byte[] value, UInt32 size)
        {
            Resize(m_wpos + size);
            Array.Copy(value, 0, m_buff, m_wpos, size);
            m_wpos += size;
        }

        void _Write(byte[] value, UInt32 offset, UInt32 size)
        {
            Resize(m_wpos + size);
            Array.Copy(value, offset, m_buff, m_wpos, size);
            m_wpos += size;
        }

        protected void _Read(byte[] dest, UInt32 size)
        {
            Array.Copy(m_buff, m_rpos, dest, 0, size);
            m_rpos += size;
        }

        public int _Read(byte[] dest, int offest, int size)
        {
            if (offest < 0)
            {
                return 0;
            }
            uint dSize = (uint)dest.Length;
            if (offest > dSize - 1)
            {
                return 0;
            }
            if (Size - 1 < m_rpos)
            {
                return 0;
            }

            UInt32 rSize = (UInt32)size;
            if (dSize - 1 < offest + rSize)
            {
                rSize = (UInt32)(dSize - offest - 1);
            }
            uint lSize = Size - m_rpos;

            if (lSize < rSize)
            {
                rSize = lSize;
            }

            if (rSize == 0)
            {
                return 0;
            }
            Array.Copy(m_buff, m_rpos, dest, offest, rSize);
            m_rpos += rSize;
            return (int)rSize;
        }

        public void Append(Byte[] src, UInt32 size) { _Write(src, size); }

        public ExcelTranslatorBuffer In(bool value)
        {
            if (value)
                __InOut_buf[0] = (byte)1;
            else
                __InOut_buf[0] = (byte)0;
            _Write(__InOut_buf, sizeof(bool));
            return this;
        }
        public ExcelTranslatorBuffer Out(out bool value)
        {
            _Read(__InOut_buf, sizeof(bool));
            value = BitConverter.ToBoolean(__InOut_buf, 0);
            return this;
        }
        public ExcelTranslatorBuffer In(bool[] valueArray)
        {
            In(valueArray.Length);
            for (int i = 0; i < valueArray.Length; i++)
            {
                In(valueArray[i]);
            }
            return this;
        }
        public ExcelTranslatorBuffer Out(out bool[] valueArray)
        {
            int length = 0;
            Out(out length);
            valueArray = new bool[length];
            for (int i = 0; i < length; i++)
            {
                bool value = true;
                Out(out value);
                valueArray[i] = value;
            }
            return this;
        }

        public ExcelTranslatorBuffer In(int value)
        {
            byte[] _bytes = BitConverter.GetBytes(value);
            _bytes.CopyTo(__InOut_buf, 0);
            _Write(__InOut_buf, sizeof(int));
            return this;
        }
        public ExcelTranslatorBuffer Out(out int value)
        {
            _Read(__InOut_buf, sizeof(int));
            value = BitConverter.ToInt32(__InOut_buf, 0);
            return this;
        }
        public ExcelTranslatorBuffer In(int[] valueArray)
        {
            In(valueArray.Length);
            for (int i = 0; i < valueArray.Length; i++)
            {
                In(valueArray[i]);
            }
            return this;
        }
        public ExcelTranslatorBuffer Out(out int[] valueArray)
        {
            int length = 0;
            Out(out length);
            valueArray = new int[length];
            for (int i = 0; i < length; i++)
            {
                int value = 0;
                Out(out value);
                valueArray[i] = value;
            }
            return this;
        }

        public ExcelTranslatorBuffer In(String value)
        {
            byte[] bytes = Encoding.UTF8.GetBytes(value);
            int size = bytes != null ? bytes.Length : 0;
            In(size);
            if (size > 0) _Write(bytes, (uint)size);
            return this;
        }
        public ExcelTranslatorBuffer Out(out String value)
        {
            int size = 0;
            Out(out size);
            if (size > 0)
            {
                _Read(__InOut_buf, (uint)size);
                value = Encoding.UTF8.GetString(__InOut_buf, 0, size);
            }
            else
            {
                value = String.Empty;
            }
            return this;
        }
        public ExcelTranslatorBuffer In(String[] valueArray)
        {
            In(valueArray.Length);
            for (int i = 0; i < valueArray.Length; i++)
            {
                In(valueArray[i]);
            }
            return this;
        }
        public ExcelTranslatorBuffer Out(out String[] valueArray)
        {
            int length = 0;
            Out(out length);
            valueArray = new String[length];
            for (int i = 0; i < length; i++)
            {
                String value = string.Empty;
                Out(out value);
                valueArray[i] = value;
            }
            return this;
        }

        public ExcelTranslatorBuffer In(float value)
        {
            byte[] _bytes = System.BitConverter.GetBytes(value);
            _bytes.CopyTo(__InOut_buf, 0);
            _Write(__InOut_buf, sizeof(float));
            return this;
        }
        public ExcelTranslatorBuffer Out(out float value)
        {
            _Read(__InOut_buf, sizeof(float));
            value = BitConverter.ToSingle(__InOut_buf, 0);
            return this;
        }
        public ExcelTranslatorBuffer In(float[] valueArray)
        {
            In(valueArray.Length);
            for (int i = 0; i < valueArray.Length; i++)
            {
                In(valueArray[i]);
            }
            return this;
        }
        public ExcelTranslatorBuffer Out(out float[] valueArray)
        {
            int length = 0;
            Out(out length);
            valueArray = new float[length];
            for (int i = 0; i < length; i++)
            {
                float value = 0f;
                Out(out value);
                valueArray[i] = value;
            }
            return this;
        }

        public ExcelTranslatorBuffer In(ValueType valueType, List<byte[]> value)
        {
            //读取数组长度
            int length = value.Count;
            if (valueType == ValueType.BoolArray ||
                valueType == ValueType.Int32Array ||
                valueType == ValueType.FloatArray ||
                valueType == ValueType.StringArray)
            {
                In(length);
            }

            for (int i = 0; i < value.Count; i++)
            {
                switch (valueType)
                {
                    case ValueType.Int32:
                    case ValueType.Int32Array:
                    case ValueType.Bool:
                    case ValueType.BoolArray:
                    case ValueType.Float:
                    case ValueType.FloatArray:
                        _Write(value[i], (uint)value[i].Length);
                        break;
                    case ValueType.String:
                    case ValueType.StringArray:
                        int size = value[i] != null ? value[i].Length : 0;
                        _Write(BitConverter.GetBytes(size),sizeof(int));
                        if (size > 0) _Write(value[i], (uint)size);
                        break;
                    default:
                        throw new Exception("ExcelTranslatorBuffer.OutDynamicValue() 不存在的类型！ " + valueType);
                }
            }
            return this;
        }
        public ExcelTranslatorBuffer Out(ValueType valueType, out List<byte[]> value)
        {
            value = new List<byte[]>();

            //读取数组长度
            int length = 1;
            if (valueType == ValueType.BoolArray ||
                valueType == ValueType.Int32Array ||
                valueType == ValueType.FloatArray ||
                valueType == ValueType.StringArray)
            {
                Out(out length);
            }

            for (int i = 0; i < length; i++)
            {
                int size = 0;
                switch (valueType)
                {
                    case ValueType.Int32:
                    case ValueType.Int32Array:
                        size = sizeof(int);
                        break;

                    case ValueType.Bool:
                    case ValueType.BoolArray:
                        size = sizeof(bool);
                        break;

                    case ValueType.Float:
                    case ValueType.FloatArray:
                        size = sizeof(float);
                        break;

                    case ValueType.String:
                    case ValueType.StringArray:
                        Out(out size);
                        break;
                    default:
                        throw new Exception("ExcelTranslatorBuffer.OutDynamicValue() 不存在的类型！ " + valueType);
                }
                byte[] bytes = new byte[size];
                if (size > 0) _Read(bytes, (uint)size);
                value.Add(bytes);
            }
            return this;
        }
    }
}
