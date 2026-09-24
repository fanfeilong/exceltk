using System;
using System.Collections.Generic;

namespace Exceltk.Reader.Binary {
    /// <summary>
    /// BIFF Package: INDEX (row block index table).
    /// </summary>
    internal class XlsBiffIndex : BinaryPackage {
        private bool isV8 = true;
        private uint m_firstExistingRow;
        private uint m_lastExistingRow;
        private uint[] m_dbCellAddresses = Array.Empty<uint>();

        internal XlsBiffIndex(ExcelBinaryReader reader)
            : base(reader) {
        }

        internal XlsBiffIndex(byte[] bytes, uint offset, ExcelBinaryReader reader)
            : base(bytes, offset, reader) {
        }

        protected override void DecodeBody(byte[] buffer, int offset, int length) {
            ParseFields(buffer, offset, length);
        }

        private void ParseFields(byte[] buffer, int offset, int length) {
            if (buffer == null || length <= 0) {
                m_dbCellAddresses = Array.Empty<uint>();
                return;
            }

            if (isV8) {
                if (length >= 12) {
                    m_firstExistingRow = BodyReadUInt32(buffer, offset, 0x4);
                    m_lastExistingRow = BodyReadUInt32(buffer, offset, 0x8);
                }
            } else {
                if (length >= 8) {
                    m_firstExistingRow = BodyReadUInt16(buffer, offset, 0x4);
                    m_lastExistingRow = BodyReadUInt16(buffer, offset, 0x6);
                }
            }

            int firstIdx = isV8 ? 16 : 12;
            if (length <= firstIdx) {
                m_dbCellAddresses = Array.Empty<uint>();
                return;
            }

            var cells = new List<uint>((length - firstIdx) / 4);
            for (int i = firstIdx; i + 4 <= length; i += 4) {
                cells.Add(BodyReadUInt32(buffer, offset, i));
            }
            m_dbCellAddresses = cells.ToArray();
        }

        protected override int GetRequiredEncodeBodyBufferLength() {
            int header = isV8 ? 16 : 12;
            return header + m_dbCellAddresses.Length * 4;
        }

        protected override void EncodeBody(byte[] buffer, int offset, int capacity, out int written) {
            int need = GetRequiredEncodeBodyBufferLength();
            if (capacity < need) {
                throw new ArgumentException(Errors.ErrorBIFFBufferSize);
            }

            // Reserved leading zeros (BIFF INDEX layout).
            for (int i = 0; i < (isV8 ? 16 : 12); i++) {
                buffer[offset + i] = 0;
            }

            if (isV8) {
                WriteUInt32(buffer, offset + 0x4, m_firstExistingRow);
                WriteUInt32(buffer, offset + 0x8, m_lastExistingRow);
                // bytes 0-3 and 12-15 remain reserved zeros
                int pos = 16;
                for (int i = 0; i < m_dbCellAddresses.Length; i++) {
                    WriteUInt32(buffer, offset + pos, m_dbCellAddresses[i]);
                    pos += 4;
                }
            } else {
                WriteUInt16(buffer, offset + 0x4, (ushort)m_firstExistingRow);
                WriteUInt16(buffer, offset + 0x6, (ushort)m_lastExistingRow);
                int pos = 12;
                for (int i = 0; i < m_dbCellAddresses.Length; i++) {
                    WriteUInt32(buffer, offset + pos, m_dbCellAddresses[i]);
                    pos += 4;
                }
            }
            written = need;
        }

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

        public uint FirstExistingRow {
            get {
                return m_firstExistingRow;
            }
        }

        public uint LastExistingRow {
            get {
                return m_lastExistingRow;
            }
        }

        public uint[] DbCellAddresses {
            get {
                return m_dbCellAddresses;
            }
        }
    }
}
