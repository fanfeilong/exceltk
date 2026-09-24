using System;
using System.Text;

namespace Exceltk.Reader.Binary {
    /// <summary>
    /// BIFF Package: FORMAT string. Eager DecodeBody fills index/length/flags/value.
    /// </summary>
    internal class XlsBiffFormatString : BinaryPackage {
        private Encoding m_UseEncoding = Exceltk.Extension.DefaultEncoding();
        private ushort m_length;
        private ushort m_index;
        private byte m_flags;
        private string m_value = string.Empty;
        private byte[] m_stringBytes = Array.Empty<byte>();
        private int m_stringBodyOfs; // relative start of character bytes within body

        internal XlsBiffFormatString(ExcelBinaryReader reader)
            : base(reader) {
        }

        internal XlsBiffFormatString(byte[] bytes, uint offset, ExcelBinaryReader reader)
            : base(bytes, offset, reader) {
        }

        protected override void DecodeBody(byte[] buffer, int offset, int length) {
            switch (ID) {
                case BIFFRECORDTYPE.FORMAT_V23:
                    if (length < 1) {
                        m_length = 0;
                        m_index = 0;
                        m_flags = 0;
                        m_stringBytes = Array.Empty<byte>();
                        m_value = string.Empty;
                        return;
                    }
                    m_length = BodyReadByte(buffer, offset, 0x0);
                    m_index = 0;
                    m_flags = 0;
                    m_stringBodyOfs = 1;
                    {
                        int n = Math.Min(m_length, Math.Max(0, length - 1));
                        m_stringBytes = new byte[n];
                        if (n > 0) {
                            Buffer.BlockCopy(buffer, offset + 1, m_stringBytes, 0, n);
                        }
                        m_value = m_UseEncoding.GetString(m_stringBytes, 0, m_stringBytes.Length);
                    }
                    break;
                default:
                    if (length < 5) {
                        m_length = 0;
                        m_index = 0;
                        m_flags = 0;
                        m_stringBytes = Array.Empty<byte>();
                        m_value = string.Empty;
                        return;
                    }
                    m_index = BodyReadUInt16(buffer, offset, 0);
                    m_length = BodyReadUInt16(buffer, offset, 2);
                    m_flags = BodyReadByte(buffer, offset, 3);
                    {
                        Encoding enc = (m_flags & 0x01) == 0x01
                            ? Encoding.Unicode
                            : Exceltk.Extension.DefaultEncoding();
                        m_UseEncoding = enc;

                        int strOfs = 5;
                        // Optional blocks (bit tests; prior ==0x01 checks were ineffective).
                        if ((m_flags & 0x04) != 0) {
                            strOfs += 4;
                        }
                        if ((m_flags & 0x08) != 0) {
                            strOfs += 2;
                        }

                        m_stringBodyOfs = strOfs;
                        int byteLen = m_UseEncoding.IsSingleByte ? m_length : m_length * 2;
                        byteLen = Math.Min(byteLen, Math.Max(0, length - strOfs));
                        m_stringBytes = new byte[byteLen];
                        if (byteLen > 0) {
                            Buffer.BlockCopy(buffer, offset + strOfs, m_stringBytes, 0, byteLen);
                        }
                        m_value = m_UseEncoding.GetString(m_stringBytes, 0, m_stringBytes.Length);
                    }
                    break;
            }
        }

        protected override int GetRequiredEncodeBodyBufferLength() {
            if (ID == BIFFRECORDTYPE.FORMAT_V23) {
                return 1 + m_stringBytes.Length;
            }
            return m_stringBodyOfs + m_stringBytes.Length;
        }

        protected override void EncodeBody(byte[] buffer, int offset, int capacity, out int written) {
            int need = GetRequiredEncodeBodyBufferLength();
            if (capacity < need) {
                throw new ArgumentException(Errors.ErrorBIFFBufferSize);
            }
            for (int i = 0; i < need; i++) {
                buffer[offset + i] = 0;
            }

            if (ID == BIFFRECORDTYPE.FORMAT_V23) {
                buffer[offset] = (byte)m_length;
                if (m_stringBytes.Length > 0) {
                    Buffer.BlockCopy(m_stringBytes, 0, buffer, offset + 1, m_stringBytes.Length);
                }
            } else {
                WriteUInt16(buffer, offset + 0, m_index);
                WriteUInt16(buffer, offset + 2, m_length);
                buffer[offset + 3] = m_flags;
                buffer[offset + 4] = 0; // reserved / high of grbit region
                if (m_stringBytes.Length > 0) {
                    Buffer.BlockCopy(m_stringBytes, 0, buffer, offset + m_stringBodyOfs, m_stringBytes.Length);
                }
            }
            written = need;
        }

        public Encoding UseEncoding {
            get {
                return m_UseEncoding;
            }
            set {
                m_UseEncoding = value;
                if (m_stringBytes.Length > 0) {
                    m_value = m_UseEncoding.GetString(m_stringBytes, 0, m_stringBytes.Length);
                }
            }
        }

        public ushort Length {
            get {
                return m_length;
            }
        }

        public string Value {
            get {
                return m_value;
            }
        }

        public ushort Index {
            get {
                return m_index;
            }
        }
    }
}
