namespace Exceltk.Reader.Binary {
    /// <summary>
    /// BIFF Package: BOOLERR cell.
    /// </summary>
    internal class XlsBiffBoolErr : XlsBiffBlankCell {
        private bool m_boolValue;

        internal XlsBiffBoolErr(ExcelBinaryReader reader)
            : base(reader) {
        }

        internal XlsBiffBoolErr(byte[] bytes, ExcelBinaryReader reader)
            : this(bytes, 0, reader) {
        }

        internal XlsBiffBoolErr(byte[] bytes, uint offset, ExcelBinaryReader reader)
            : base(bytes, offset, reader) {
        }

        protected override void DecodeBody(byte[] buffer, int offset, int length) {
            base.DecodeBody(buffer, offset, length);
            m_boolValue = ReadByte(0x6) == 1;
        }

        protected override int GetRequiredEncodeBodyBufferLength() {
            return 8;
        }

        protected override void EncodeBody(byte[] buffer, int offset, int capacity, out int written) {
            base.EncodeBody(buffer, offset, capacity, out written);
            if (capacity < 8) throw new System.ArgumentException(Errors.ErrorBIFFBufferSize);
            buffer[offset + 0x6] = (byte)(m_boolValue ? 1 : 0);
            buffer[offset + 0x7] = 0;
            written = 8;
        }

        public bool BoolValue {
            get {
                return m_boolValue;
            }
        }
    }
}
