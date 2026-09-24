namespace Exceltk.Reader.Binary {
    /// <summary>
    /// Represents a constant integer number in range 0..65535
    /// </summary>
    internal class XlsBiffIntegerCell : XlsBiffBlankCell {
        private uint m_value;

        internal XlsBiffIntegerCell(ExcelBinaryReader reader)
            : base(reader) {
        }

        internal XlsBiffIntegerCell(byte[] bytes, uint offset, ExcelBinaryReader reader)
            : base(bytes, offset, reader) {
        }

        protected override void DecodeBody(byte[] buffer, int offset, int length) {
            base.DecodeBody(buffer, offset, length);
            m_value = ReadUInt16(0x6);
        }

        protected override int GetRequiredEncodeBodyBufferLength() {
            return 8;
        }

        protected override void EncodeBody(byte[] buffer, int offset, int capacity, out int written) {
            base.EncodeBody(buffer, offset, capacity, out written);
            if (capacity < 8) throw new System.ArgumentException(Errors.ErrorBIFFBufferSize);
            WriteUInt16(buffer, offset + 0x6, (ushort)m_value);
            written = 8;
        }

        /// <summary>
        /// Returns value of this cell
        /// </summary>
        public uint Value {
            get {
                return m_value;
            }
        }
    }
}
