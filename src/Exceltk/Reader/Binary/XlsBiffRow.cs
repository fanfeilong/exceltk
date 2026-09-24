namespace Exceltk.Reader.Binary {
    /// <summary>
    /// BIFF Package: ROW record.
    /// </summary>
    internal class XlsBiffRow : BinaryPackage {
        private ushort m_rowIndex;
        private ushort m_firstDefinedColumn;
        private ushort m_lastDefinedColumn;
        private ushort m_rowHeight;
        private ushort m_flags;
        private ushort m_xFormat;

        internal XlsBiffRow(ExcelBinaryReader reader)
            : base(reader) {
        }

        internal XlsBiffRow(byte[] bytes, uint offset, ExcelBinaryReader reader)
            : base(bytes, offset, reader) {
        }

        protected override void DecodeBody(byte[] buffer, int offset, int length) {
            m_rowIndex = ReadUInt16(0x0);
            m_firstDefinedColumn = ReadUInt16(0x2);
            m_lastDefinedColumn = ReadUInt16(0x4);
            m_rowHeight = ReadUInt16(0x6);
            m_flags = ReadUInt16(0xC);
            m_xFormat = ReadUInt16(0xE);
        }

        protected override int GetRequiredEncodeBodyBufferLength() {
            return 16;
        }

        protected override void EncodeBody(byte[] buffer, int offset, int capacity, out int written) {
            if (capacity < 16) throw new System.ArgumentException(Errors.ErrorBIFFBufferSize);
            WriteUInt16(buffer, offset + 0x0, m_rowIndex);
            WriteUInt16(buffer, offset + 0x2, m_firstDefinedColumn);
            WriteUInt16(buffer, offset + 0x4, m_lastDefinedColumn);
            WriteUInt16(buffer, offset + 0x6, m_rowHeight);
            WriteUInt16(buffer, offset + 0x8, 0);
            WriteUInt16(buffer, offset + 0xA, 0);
            WriteUInt16(buffer, offset + 0xC, m_flags);
            WriteUInt16(buffer, offset + 0xE, m_xFormat);
            written = 16;
        }

        /// <summary>
        /// Zero-based index of row described
        /// </summary>
        public ushort RowIndex {
            get {
                return m_rowIndex;
            }
        }

        /// <summary>
        /// Index of first defined column
        /// </summary>
        public ushort FirstDefinedColumn {
            get {
                return m_firstDefinedColumn;
            }
        }

        /// <summary>
        /// Index of last defined column
        /// </summary>
        public ushort LastDefinedColumn {
            get {
                return m_lastDefinedColumn;
            }
        }

        /// <summary>
        /// Returns row height
        /// </summary>
        public uint RowHeight {
            get {
                return m_rowHeight;
            }
        }

        /// <summary>
        /// Returns row flags
        /// </summary>
        public ushort Flags {
            get {
                return m_flags;
            }
        }

        /// <summary>
        /// Returns default format for this row
        /// </summary>
        public ushort XFormat {
            get {
                return m_xFormat;
            }
        }
    }
}
