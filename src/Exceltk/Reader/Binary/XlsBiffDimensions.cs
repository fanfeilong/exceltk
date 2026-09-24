using System;

namespace Exceltk.Reader.Binary {
    /// <summary>
    /// BIFF Package: DIMENSIONS.
    /// </summary>
    internal class XlsBiffDimensions : BinaryPackage {
        private bool isV8 = true;
        private uint m_firstRow;
        private uint m_lastRow;
        private ushort m_firstColumn;
        private ushort m_lastColumn;

        internal XlsBiffDimensions(ExcelBinaryReader reader)
            : base(reader) {
        }

        internal XlsBiffDimensions(byte[] bytes, uint offset, ExcelBinaryReader reader)
            : base(bytes, offset, reader) {
        }

        protected override void DecodeBody(byte[] buffer, int offset, int length) {
            ParseFields(buffer, offset, length);
        }

        private void ParseFields(byte[] buffer, int offset, int length) {
            if (buffer == null || length <= 0) {
                return;
            }
            if (isV8) {
                if (length < 10) {
                    return;
                }
                m_firstRow = BodyReadUInt32(buffer, offset, 0x0);
                m_lastRow = BodyReadUInt32(buffer, offset, 0x4);
                m_firstColumn = BodyReadUInt16(buffer, offset, 0x8);
                // Preserve historical decode: high byte of UInt16 at body+0x9, then +1.
                m_lastColumn = (ushort)((BodyReadUInt16(buffer, offset, 0x9) >> 8) + 1);
            } else {
                if (length < 8) {
                    return;
                }
                m_firstRow = BodyReadUInt16(buffer, offset, 0x0);
                m_lastRow = BodyReadUInt16(buffer, offset, 0x2);
                m_firstColumn = BodyReadUInt16(buffer, offset, 0x4);
                m_lastColumn = BodyReadUInt16(buffer, offset, 0x6);
            }
        }

        protected override int GetRequiredEncodeBodyBufferLength() {
            return isV8 ? 14 : 8;
        }

        protected override void EncodeBody(byte[] buffer, int offset, int capacity, out int written) {
            int need = GetRequiredEncodeBodyBufferLength();
            if (capacity < need) {
                throw new ArgumentException(Errors.ErrorBIFFBufferSize);
            }
            if (isV8) {
                WriteUInt32(buffer, offset + 0x0, m_firstRow);
                WriteUInt32(buffer, offset + 0x4, m_lastRow);
                WriteUInt16(buffer, offset + 0x8, m_firstColumn);
                // Inverse of decode (ReadUInt16(0x9)>>8)+1 → low byte of colMac at 0xA.
                ushort colMacWire = m_lastColumn == 0 ? (ushort)0 : (ushort)(m_lastColumn - 1);
                WriteUInt16(buffer, offset + 0xA, colMacWire);
                WriteUInt16(buffer, offset + 0xC, 0); // reserved
            } else {
                WriteUInt16(buffer, offset + 0x0, (ushort)m_firstRow);
                WriteUInt16(buffer, offset + 0x2, (ushort)m_lastRow);
                WriteUInt16(buffer, offset + 0x4, m_firstColumn);
                WriteUInt16(buffer, offset + 0x6, m_lastColumn);
            }
            written = need;
        }

        /// <summary>
        /// Gets or sets if BIFF8 addressing is used
        /// </summary>
        public bool IsV8 {
            get {
                return isV8;
            }
            set {
                isV8 = value;
                if (m_bytes != null && m_bytes.Length > 0) {
                    ParseFields(m_bytes, m_readoffset, m_bodyLength);
                }
            }
        }

        public uint FirstRow {
            get {
                return m_firstRow;
            }
        }

        public uint LastRow {
            get {
                return m_lastRow;
            }
        }

        public ushort FirstColumn {
            get {
                return m_firstColumn;
            }
        }

        public ushort LastColumn {
            get {
                return m_lastColumn;
            }
            set {
                throw new NotImplementedException();
            }
        }
    }
}
