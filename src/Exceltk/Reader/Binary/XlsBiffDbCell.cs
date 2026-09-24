using System;
using System.Collections.Generic;

namespace Exceltk.Reader.Binary {
    /// <summary>
    /// BIFF Package: DBCELL (row block index).
    /// Wire: Int32 at 0 (offset to first ROW), then rgdb as UInt16 every 4 bytes (2-byte pad).
    /// </summary>
    internal class XlsBiffDbCell : BinaryPackage {
        /// <summary>Wire Int32 at body offset 0.</summary>
        private int m_rowOffsetWire;
        /// <summary>Wire rgdb ushorts (one every 4 bytes starting at body+4).</summary>
        private ushort[] m_rgdb = Array.Empty<ushort>();

        private int m_rowAddress;
        private uint[] m_cellAddresses = Array.Empty<uint>();

        internal XlsBiffDbCell(ExcelBinaryReader reader)
            : base(reader) {
        }

        internal XlsBiffDbCell(byte[] bytes, uint offset, ExcelBinaryReader reader)
            : base(bytes, offset, reader) {
        }

        protected override void DecodeBody(byte[] buffer, int offset, int length) {
            if (length < 4) {
                m_rowOffsetWire = 0;
                m_rgdb = Array.Empty<ushort>();
                m_rowAddress = Offset;
                m_cellAddresses = Array.Empty<uint>();
                return;
            }

            m_rowOffsetWire = BodyReadInt32(buffer, offset, 0x0);
            m_rowAddress = Offset - m_rowOffsetWire;

            var rgdb = new List<ushort>();
            for (int i = 0x4; i + 2 <= length; i += 4) {
                rgdb.Add(BodyReadUInt16(buffer, offset, i));
            }
            m_rgdb = rgdb.ToArray();

            int a = m_rowAddress - 20; // 20 assumed to be row structure size
            var tmp = new List<uint>(m_rgdb.Length);
            for (int i = 0; i < m_rgdb.Length; i++) {
                tmp.Add((uint)a + m_rgdb[i]);
            }
            m_cellAddresses = tmp.ToArray();
        }

        protected override int GetRequiredEncodeBodyBufferLength() {
            return 4 + m_rgdb.Length * 4;
        }

        protected override void EncodeBody(byte[] buffer, int offset, int capacity, out int written) {
            int need = GetRequiredEncodeBodyBufferLength();
            if (capacity < need) {
                throw new ArgumentException(Errors.ErrorBIFFBufferSize);
            }
            WriteInt32(buffer, offset + 0x0, m_rowOffsetWire);
            int pos = 4;
            for (int i = 0; i < m_rgdb.Length; i++) {
                WriteUInt16(buffer, offset + pos, m_rgdb[i]);
                buffer[offset + pos + 2] = 0;
                buffer[offset + pos + 3] = 0;
                pos += 4;
            }
            written = need;
        }

        /// <summary>Offset of first row linked with this record</summary>
        public int RowAddress {
            get {
                return m_rowAddress;
            }
        }

        /// <summary>Addresses of cell values</summary>
        public uint[] CellAddresses {
            get {
                return m_cellAddresses;
            }
        }
    }
}
