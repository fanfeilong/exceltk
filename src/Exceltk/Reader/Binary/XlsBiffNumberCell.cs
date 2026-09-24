namespace Exceltk.Reader.Binary {
    /// <summary>
    /// BIFF Package: NUMBER cell (IEEE double).
    /// </summary>
    internal class XlsBiffNumberCell : XlsBiffBlankCell {
        private double m_value;

        internal XlsBiffNumberCell(ExcelBinaryReader reader)
            : base(reader) {
        }

        internal XlsBiffNumberCell(byte[] bytes, uint offset, ExcelBinaryReader reader)
            : base(bytes, offset, reader) {
        }

        protected override void DecodeBody(byte[] buffer, int offset, int length) {
            base.DecodeBody(buffer, offset, length);
            m_value = ReadDouble(0x6);
        }

        protected override int GetRequiredEncodeBodyBufferLength() {
            return 14;
        }

        protected override void EncodeBody(byte[] buffer, int offset, int capacity, out int written) {
            base.EncodeBody(buffer, offset, capacity, out written);
            if (capacity < 14) throw new System.ArgumentException(Errors.ErrorBIFFBufferSize);
            WriteDouble(buffer, offset + 0x6, m_value);
            written = 14;
        }

        public double Value {
            get {
                return m_value;
            }
        }
    }
}
