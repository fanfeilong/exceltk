using System;
using System.Text;

namespace Exceltk.Reader.Binary {
    /// <summary>
    /// BIFF Package: BOUNDSHEET (sheet metadata in workbook globals).
    /// </summary>
    internal class XlsBiffBoundSheet : BinaryPackage {
        #region SheetType enum

        public enum SheetType : byte {
            Worksheet = 0x0,
            MacroSheet = 0x1,
            Chart = 0x2,
            VBModule = 0x6
        }

        #endregion

        #region SheetVisibility enum

        public enum SheetVisibility : byte {
            Visible = 0x0,
            Hidden = 0x1,
            VeryHidden = 0x2
        }

        #endregion

        private bool isV8 = true;
        private Encoding m_UseEncoding = Exceltk.Extension.DefaultEncoding();
        private uint m_startOffset;
        private SheetType m_type;
        private SheetVisibility m_visibleState;
        private byte m_nameCharCount;
        private byte m_nameFlags; // BIFF8 option flags at body+0x7 (0 = compressed/default)
        private string m_sheetName = string.Empty;

        internal XlsBiffBoundSheet(ExcelBinaryReader reader)
            : base(reader) {
        }

        internal XlsBiffBoundSheet(byte[] bytes, uint offset, ExcelBinaryReader reader)
            : base(bytes, offset, reader) {
        }

        protected override void DecodeBody(byte[] buffer, int offset, int length) {
            if (length < 6) {
                return;
            }
            m_startOffset = BodyReadUInt32(buffer, offset, 0x0);
            m_type = (SheetType)BodyReadByte(buffer, offset, 0x4);
            m_visibleState = (SheetVisibility)(BodyReadByte(buffer, offset, 0x5) & 0x3);
            ParseSheetName(buffer, offset, length);
        }

        private void ParseSheetName(byte[] buffer, int offset, int length) {
            if (length < 7) {
                m_sheetName = string.Empty;
                return;
            }
            m_nameCharCount = BodyReadByte(buffer, offset, 0x6);
            if (isV8) {
                m_nameFlags = length > 7 ? BodyReadByte(buffer, offset, 0x7) : (byte)0;
                const int start = 0x8;
                int chars = m_nameCharCount;
                if (m_nameFlags == 0) {
                    int byteLen = Math.Min(chars, Math.Max(0, length - start));
                    m_sheetName = Exceltk.Extension.DefaultEncoding().GetString(buffer, offset + start, byteLen);
                } else {
                    int byteLen = m_UseEncoding.IsSingleByteEncoding() ? chars : chars * 2;
                    byteLen = Math.Min(byteLen, Math.Max(0, length - start));
                    m_sheetName = m_UseEncoding.GetString(buffer, offset + start, byteLen);
                }
            } else {
                m_nameFlags = 0;
                const int start = 0x7;
                int byteLen = Math.Min(m_nameCharCount, Math.Max(0, length - start));
                m_sheetName = Exceltk.Extension.DefaultEncoding().GetString(buffer, offset + start, byteLen);
            }
        }

        protected override int GetRequiredEncodeBodyBufferLength() {
            if (isV8) {
                int nameBytes = m_nameFlags == 0
                    ? m_nameCharCount
                    : (m_UseEncoding.IsSingleByteEncoding() ? m_nameCharCount : m_nameCharCount * 2);
                return 8 + nameBytes;
            }
            return 7 + m_nameCharCount;
        }

        protected override void EncodeBody(byte[] buffer, int offset, int capacity, out int written) {
            int need = GetRequiredEncodeBodyBufferLength();
            if (capacity < need) {
                throw new ArgumentException(Errors.ErrorBIFFBufferSize);
            }
            WriteUInt32(buffer, offset + 0x0, m_startOffset);
            buffer[offset + 0x4] = (byte)m_type;
            buffer[offset + 0x5] = (byte)m_visibleState;
            buffer[offset + 0x6] = m_nameCharCount;

            byte[] nameBytes;
            if (isV8) {
                buffer[offset + 0x7] = m_nameFlags;
                if (m_nameFlags == 0) {
                    nameBytes = Exceltk.Extension.DefaultEncoding().GetBytes(m_sheetName ?? string.Empty);
                } else {
                    nameBytes = m_UseEncoding.GetBytes(m_sheetName ?? string.Empty);
                }
                int copy = Math.Min(nameBytes.Length, need - 8);
                if (copy > 0) {
                    Buffer.BlockCopy(nameBytes, 0, buffer, offset + 0x8, copy);
                }
            } else {
                nameBytes = Exceltk.Extension.DefaultEncoding().GetBytes(m_sheetName ?? string.Empty);
                int copy = Math.Min(nameBytes.Length, need - 7);
                if (copy > 0) {
                    Buffer.BlockCopy(nameBytes, 0, buffer, offset + 0x7, copy);
                }
            }
            written = need;
        }

        public uint StartOffset {
            get {
                return m_startOffset;
            }
        }

        public SheetType Type {
            get {
                return m_type;
            }
        }

        public SheetVisibility VisibleState {
            get {
                return m_visibleState;
            }
        }

        public string SheetName {
            get {
                return m_sheetName;
            }
        }

        public Encoding UseEncoding {
            get {
                return m_UseEncoding;
            }
            set {
                m_UseEncoding = value;
                if (m_bytes != null && m_bytes.Length > 0) {
                    ParseSheetName(m_bytes, m_readoffset, m_bodyLength);
                }
            }
        }

        public bool IsV8 {
            get {
                return isV8;
            }
            set {
                isV8 = value;
                if (m_bytes != null && m_bytes.Length > 0) {
                    ParseSheetName(m_bytes, m_readoffset, m_bodyLength);
                }
            }
        }
    }
}
