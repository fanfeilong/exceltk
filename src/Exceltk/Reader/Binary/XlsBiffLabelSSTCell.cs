namespace Exceltk.Reader.Binary {
    /// <summary>
    /// BIFF Package: label cell referencing SST. Body after blank header: SST index at 0x6.
    /// </summary>
    internal class XlsBiffLabelSSTCell : XlsBiffBlankCell {
        private uint m_sstIndex;

        internal XlsBiffLabelSSTCell(ExcelBinaryReader reader)
            : base(reader) {
        }

        internal XlsBiffLabelSSTCell(byte[] bytes, uint offset, ExcelBinaryReader reader)
            : base(bytes, offset, reader) {
        }

        protected override void DecodeBody(byte[] buffer, int offset, int length) {
            base.DecodeBody(buffer, offset, length);
            m_sstIndex = ReadUInt32(0x6);
        }

        protected override int GetRequiredEncodeBodyBufferLength() {
            return 10;
        }

        protected override void EncodeBody(byte[] buffer, int offset, int capacity, out int written) {
            base.EncodeBody(buffer, offset, capacity, out written);
            if (capacity < 10) throw new System.ArgumentException(Errors.ErrorBIFFBufferSize);
            WriteUInt32(buffer, offset + 0x6, m_sstIndex);
            written = 10;
        }

        public uint SSTIndex {
            get {
                return m_sstIndex;
            }
        }

        public string Text(XlsBiffSST sst) {
            return sst.GetString(SSTIndex);
        }
    }
}
