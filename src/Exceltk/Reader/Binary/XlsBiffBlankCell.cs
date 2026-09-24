namespace Exceltk.Reader.Binary {
    /// <summary>
    /// BIFF Package mid-base for cell records (BLANK and shared row/col/xf prefix).
    /// Concrete cell Packages (RK, Number, LabelSST, …) inherit this.
    /// </summary>
    internal class XlsBiffBlankCell : BinaryPackage {
        private ushort m_rowIndex;
        private ushort m_columnIndex;
        private ushort m_xFormat;

        internal XlsBiffBlankCell(ExcelBinaryReader reader)
            : base(reader) {
        }

        internal XlsBiffBlankCell(byte[] bytes, uint offset, ExcelBinaryReader reader)
            : base(bytes, offset, reader) {
        }

        protected override void DecodeBody(byte[] buffer, int offset, int length) {
            m_rowIndex = ReadUInt16(0x0);
            m_columnIndex = ReadUInt16(0x2);
            m_xFormat = ReadUInt16(0x4);
        }

        protected override int GetRequiredEncodeBodyBufferLength() {
            return 6;
        }

        protected override void EncodeBody(byte[] buffer, int offset, int capacity, out int written) {
            if (capacity < 6) throw new System.ArgumentException(Errors.ErrorBIFFBufferSize);
            WriteUInt16(buffer, offset + 0x0, m_rowIndex);
            WriteUInt16(buffer, offset + 0x2, m_columnIndex);
            WriteUInt16(buffer, offset + 0x4, m_xFormat);
            written = 6;
        }

        /// <summary>Zero-based index of row containing this cell</summary>
        public ushort RowIndex {
            get {
                return m_rowIndex;
            }
        }

        /// <summary>Zero-based index of column containing this cell</summary>
        public ushort ColumnIndex {
            get {
                return m_columnIndex;
            }
        }

        /// <summary>Format used for this cell</summary>
        public ushort XFormat {
            get {
                return m_xFormat;
            }
        }
    }
}
