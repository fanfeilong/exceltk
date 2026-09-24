using System;
using System.Text;

namespace Exceltk.Reader.Binary {
    /// <summary>
    /// BIFF Package: LABEL / RSTRING cell (inline string).
    /// </summary>
    internal class XlsBiffLabelCell : XlsBiffBlankCell {
        private Encoding m_UseEncoding = Exceltk.Extension.DefaultEncoding();
        private ushort m_length;
        private bool m_isV8;
        private byte m_biff8Flags;
        private string m_value = string.Empty;
        private byte[] m_stringBytes = Array.Empty<byte>();

        internal XlsBiffLabelCell(ExcelBinaryReader reader)
            : base(reader) {
            m_isV8 = reader != null && reader.isV8();
        }

        internal XlsBiffLabelCell(byte[] bytes, uint offset, ExcelBinaryReader reader)
            : base(bytes, offset, reader) {
            m_isV8 = reader != null && reader.isV8();
        }

        protected override void DecodeBody(byte[] buffer, int offset, int length) {
            base.DecodeBody(buffer, offset, length);
            m_isV8 = reader != null && reader.isV8();
            if (length < 8) {
                m_length = 0;
                m_value = string.Empty;
                m_stringBytes = Array.Empty<byte>();
                return;
            }

            m_length = BodyReadUInt16(buffer, offset, 0x6);
            int unit = m_UseEncoding.IsSingleByteEncoding() ? 1 : 2;
            int charBytes = m_length * unit;

            int start;
            if (m_isV8) {
                m_biff8Flags = length > 8 ? BodyReadByte(buffer, offset, 0x8) : (byte)0;
                start = 0x9; // issue 11636
            } else {
                m_biff8Flags = 0;
                // Preserve prior BIFF3-5 reader offset for character data.
                start = 0x2;
            }

            charBytes = Math.Min(charBytes, Math.Max(0, length - start));
            m_stringBytes = new byte[charBytes];
            if (charBytes > 0) {
                Buffer.BlockCopy(buffer, offset + start, m_stringBytes, 0, charBytes);
            }
            m_value = m_UseEncoding.GetString(m_stringBytes, 0, m_stringBytes.Length);
        }

        protected override int GetRequiredEncodeBodyBufferLength() {
            if (m_isV8) {
                return 9 + m_stringBytes.Length;
            }
            // Encode BIFF3-5 with Office layout: prefix 8 + string (string stored as decoded bytes).
            return 8 + m_stringBytes.Length;
        }

        protected override void EncodeBody(byte[] buffer, int offset, int capacity, out int written) {
            int need = GetRequiredEncodeBodyBufferLength();
            if (capacity < need) {
                throw new ArgumentException(Errors.ErrorBIFFBufferSize);
            }
            WriteUInt16(buffer, offset + 0x0, RowIndex);
            WriteUInt16(buffer, offset + 0x2, ColumnIndex);
            WriteUInt16(buffer, offset + 0x4, XFormat);
            WriteUInt16(buffer, offset + 0x6, m_length);
            if (m_isV8) {
                buffer[offset + 0x8] = m_biff8Flags;
                if (m_stringBytes.Length > 0) {
                    Buffer.BlockCopy(m_stringBytes, 0, buffer, offset + 0x9, m_stringBytes.Length);
                }
            } else {
                if (m_stringBytes.Length > 0) {
                    Buffer.BlockCopy(m_stringBytes, 0, buffer, offset + 0x8, m_stringBytes.Length);
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
    }
}
