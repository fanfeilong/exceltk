namespace Exceltk.Reader.Binary {
    /// <summary>
    /// BIFF Package: EOF (end of stream section). Header-only; no body fields.
    /// </summary>
    internal class XlsBiffEOF : BinaryPackage {
        internal XlsBiffEOF(ExcelBinaryReader reader)
            : base(reader) {
        }

        internal XlsBiffEOF(byte[] bytes, uint offset, ExcelBinaryReader reader)
            : base(bytes, offset, reader) {
        }

        protected override void DecodeBody(byte[] buffer, int offset, int length) {
            // Header-only / no scalar Office fields.
        }

        protected override int GetRequiredEncodeBodyBufferLength() {
            return 0;
        }

        protected override void EncodeBody(byte[] buffer, int offset, int capacity, out int written) {
            written = 0;
        }
    }
}
