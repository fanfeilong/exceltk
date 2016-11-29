namespace Exceltk.Reader.Binary {
    internal class XlsBiffBoolErr : XlsBiffBlankCell {
        internal XlsBiffBoolErr(byte[] bytes, ExcelBinaryReader reader)
            : this(bytes, 0, reader) {
        }

        internal XlsBiffBoolErr(byte[] bytes, uint offset, ExcelBinaryReader reader)
            : base(bytes, offset, reader) {
        }

        public bool BoolValue {
            get {
                return ReadByte(0x6)==1;
            }
        }
    }
}