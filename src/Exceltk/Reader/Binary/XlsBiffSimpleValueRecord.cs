namespace Exceltk.Reader.Binary {
    /// <summary>
    /// Represents record with the only two-bytes value
    /// </summary>
    internal class XlsBiffSimpleValueRecord : BinaryPackage {
        private ushort m_value;

        internal XlsBiffSimpleValueRecord(ExcelBinaryReader reader)
            : base(reader) {
        }

        internal XlsBiffSimpleValueRecord(byte[] bytes, uint offset, ExcelBinaryReader reader)
            : base(bytes, offset, reader) {
        }

        protected override void DecodeBody(byte[] buffer, int offset, int length) {
            m_value = ReadUInt16(0x0);
        }

        protected override int GetRequiredEncodeBodyBufferLength() {
            return 2;
        }

        protected override void EncodeBody(byte[] buffer, int offset, int capacity, out int written) {
            if (capacity < 2) throw new System.ArgumentException(Errors.ErrorBIFFBufferSize);
            WriteUInt16(buffer, offset, m_value);
            written = 2;
        }

        /// <summary>
        /// Returns value
        /// </summary>
        public ushort Value {
            get {
                return m_value;
            }
        }
    }
}
