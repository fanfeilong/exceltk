using System;
using System.Text;

namespace Exceltk.Reader.Binary {
    /// <summary>
    /// BIFF Package: STRING (formula string result).
    /// </summary>
    internal class XlsBiffFormulaString : BinaryPackage {
        private const int LEADING_BYTES_COUNT = 3;
        private Encoding m_UseEncoding = Exceltk.Extension.DefaultEncoding();
        private ushort m_length;
        private bool m_isUnicode;
        private string m_value = string.Empty;
        private byte[] m_stringBytes = Array.Empty<byte>();

        internal XlsBiffFormulaString(ExcelBinaryReader reader)
            : base(reader) {
        }

        internal XlsBiffFormulaString(byte[] bytes, uint offset, ExcelBinaryReader reader)
            : base(bytes, offset, reader) {
        }

        protected override void DecodeBody(byte[] buffer, int offset, int length) {
            if (length < LEADING_BYTES_COUNT) {
                m_length = 0;
                m_isUnicode = false;
                m_value = string.Empty;
                m_stringBytes = Array.Empty<byte>();
                return;
            }

            m_length = BodyReadUInt16(buffer, offset, 0x0);
            // Preserve historical flag read (UInt16 at +1 overlaps length high byte).
            m_isUnicode = BodyReadUInt16(buffer, offset, 0x01) != 0;

            int byteLen = m_isUnicode ? m_length * 2 : m_length;
            byteLen = Math.Min(byteLen, Math.Max(0, length - LEADING_BYTES_COUNT));
            m_stringBytes = new byte[byteLen];
            if (byteLen > 0) {
                Buffer.BlockCopy(buffer, offset + LEADING_BYTES_COUNT, m_stringBytes, 0, byteLen);
            }

            if (m_isUnicode) {
                m_value = Encoding.Unicode.GetString(m_stringBytes, 0, m_stringBytes.Length);
            } else {
                m_value = m_UseEncoding.GetString(m_stringBytes, 0, m_stringBytes.Length);
            }
        }

        protected override int GetRequiredEncodeBodyBufferLength() {
            return LEADING_BYTES_COUNT + m_stringBytes.Length;
        }

        protected override void EncodeBody(byte[] buffer, int offset, int capacity, out int written) {
            int need = GetRequiredEncodeBodyBufferLength();
            if (capacity < need) {
                throw new ArgumentException(Errors.ErrorBIFFBufferSize);
            }
            WriteUInt16(buffer, offset + 0x0, m_length);
            buffer[offset + 0x2] = (byte)(m_isUnicode ? 1 : 0);
            if (m_stringBytes.Length > 0) {
                Buffer.BlockCopy(m_stringBytes, 0, buffer, offset + LEADING_BYTES_COUNT, m_stringBytes.Length);
            }
            written = need;
        }

        public Encoding UseEncoding {
            get {
                return m_UseEncoding;
            }
            set {
                m_UseEncoding = value;
                if (!m_isUnicode && m_stringBytes.Length > 0) {
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
