namespace Exceltk.Reader.Binary {
    /// <summary>
    /// BIFF Package: UNCALCED marker (header-only; no body fields).
    /// If present the Calculate Message was in the status bar when Excel saved the file.
    /// </summary>
    internal class XlsBiffUncalced : BinaryPackage {
        internal XlsBiffUncalced(ExcelBinaryReader reader)
            : base(reader) {
        }

        internal XlsBiffUncalced(byte[] bytes, uint offset, ExcelBinaryReader reader)
            : base(bytes, offset, reader) {
        }

        protected override void DecodeBody(byte[] buffer, int offset, int length) {
        }

        protected override int GetRequiredEncodeBodyBufferLength() {
            return 0;
        }

        protected override void EncodeBody(byte[] buffer, int offset, int capacity, out int written) {
            written = 0;
        }
    }
}
