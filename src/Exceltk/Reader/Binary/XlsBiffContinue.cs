using System;

namespace Exceltk.Reader.Binary {
    /// <summary>
    /// BIFF Package: CONTINUE (overflow body for large records such as SST).
    /// Opaque payload member <see cref="m_payload"/>.
    /// </summary>
    internal class XlsBiffContinue : BinaryPackage {
        private byte[] m_payload = Array.Empty<byte>();

        internal XlsBiffContinue(ExcelBinaryReader reader)
            : base(reader) {
        }

        internal XlsBiffContinue(byte[] bytes, uint offset, ExcelBinaryReader reader)
            : base(bytes, offset, reader) {
        }

        protected override void DecodeBody(byte[] buffer, int offset, int length) {
            m_payload = new byte[length];
            if (length > 0) {
                Buffer.BlockCopy(buffer, offset, m_payload, 0, length);
            }
        }

        protected override int GetRequiredEncodeBodyBufferLength() {
            return m_payload.Length;
        }

        protected override void EncodeBody(byte[] buffer, int offset, int capacity, out int written) {
            if (capacity < m_payload.Length) {
                throw new ArgumentException(Errors.ErrorBIFFBufferSize);
            }
            if (m_payload.Length > 0) {
                Buffer.BlockCopy(m_payload, 0, buffer, offset, m_payload.Length);
            }
            written = m_payload.Length;
        }

        internal byte[] Payload {
            get {
                return m_payload;
            }
        }
    }
}
