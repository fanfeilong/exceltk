namespace Exceltk.Reader.Binary {
    /// <summary>
    /// BIFF Package: RK number cell. Body after blank header: RK-encoded uint32 at 0x6.
    /// </summary>
    internal class XlsBiffRKCell : XlsBiffBlankCell {
        private double m_value;
        private uint m_rk;

        internal XlsBiffRKCell(ExcelBinaryReader reader)
            : base(reader) {
        }

        internal XlsBiffRKCell(byte[] bytes, uint offset, ExcelBinaryReader reader)
            : base(bytes, offset, reader) {
        }

        protected override void DecodeBody(byte[] buffer, int offset, int length) {
            base.DecodeBody(buffer, offset, length);
            m_rk = ReadUInt32(0x6);
            m_value = NumFromRK(m_rk);
        }

        protected override int GetRequiredEncodeBodyBufferLength() {
            return 10;
        }

        protected override void EncodeBody(byte[] buffer, int offset, int capacity, out int written) {
            base.EncodeBody(buffer, offset, capacity, out written);
            if (capacity < 10) {
                throw new System.ArgumentException(Errors.ErrorBIFFBufferSize);
            }
            WriteUInt32(buffer, offset + 0x6, m_rk);
            written = 10;
        }

        public double Value {
            get {
                return m_value;
            }
        }

        /// <summary>Decodes RK-encoded number (BIFF)</summary>
        public static double NumFromRK(uint rk) {
            double num;

            if ((rk & 0x2) == 0x2) {
                num = (int)(rk >> 2 | ((rk & 0x80000000) == 0 ? 0 : 0xC0000000));
            } else {
                long v = ((long)(rk & 0xfffffffc) << 32);
                num = v.Int64BitsToDouble();
            }

            if ((rk & 0x1) == 0x1) {
                num /= 100;
            }

            return num;
        }
    }
}
