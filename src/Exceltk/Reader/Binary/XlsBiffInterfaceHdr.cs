namespace Exceltk.Reader.Binary {
    /// <summary>
    /// Represents InterfaceHdr record in Wokrbook Globals
    /// </summary>
    internal class XlsBiffInterfaceHdr : BinaryPackage {
        private ushort m_codePage;

        internal XlsBiffInterfaceHdr(ExcelBinaryReader reader)
            : base(reader) {
        }

        internal XlsBiffInterfaceHdr(byte[] bytes, uint offset, ExcelBinaryReader reader)
            : base(bytes, offset, reader) {
        }

        protected override void DecodeBody(byte[] buffer, int offset, int length) {
            m_codePage = ReadUInt16(0x0);
        }

        protected override int GetRequiredEncodeBodyBufferLength() {
            return 2;
        }

        protected override void EncodeBody(byte[] buffer, int offset, int capacity, out int written) {
            if (capacity < 2) throw new System.ArgumentException(Errors.ErrorBIFFBufferSize);
            WriteUInt16(buffer, offset, m_codePage);
            written = 2;
        }

        /// <summary>
        /// Returns CodePage for Interface Header
        /// </summary>
        public ushort CodePage {
            get {
                return m_codePage;
            }
        }
    }
}
