using System;
using System.Text;

namespace Exceltk.Reader.Binary {
    /// <summary>
    /// BIFF Package: FORMULA cell (extends blank prefix; result is Int64 bits, not NUMBER double layout).
    /// </summary>
    internal class XlsBiffFormulaCell : XlsBiffBlankCell {
        #region FormulaFlags enum

        [Flags]
        public enum FormulaFlags : ushort {
            AlwaysCalc = 0x0001,
            CalcOnLoad = 0x0002,
            SharedFormulaGroup = 0x0008
        }

        #endregion

        private Encoding m_UseEncoding = Exceltk.Extension.DefaultEncoding();
        private FormulaFlags m_flags;
        private byte m_formulaLength;
        private long m_rawVal;
        private byte[] m_formulaBytes = Array.Empty<byte>();
        private string m_formula = string.Empty;

        internal XlsBiffFormulaCell(ExcelBinaryReader reader)
            : base(reader) {
        }

        internal XlsBiffFormulaCell(byte[] bytes, uint offset, ExcelBinaryReader reader)
            : base(bytes, offset, reader) {
        }

        protected override void DecodeBody(byte[] buffer, int offset, int length) {
            base.DecodeBody(buffer, offset, length);
            if (length < 0x10) {
                m_rawVal = 0;
                m_flags = 0;
                m_formulaLength = 0;
                m_formulaBytes = Array.Empty<byte>();
                m_formula = string.Empty;
                return;
            }

            m_rawVal = BodyReadInt64(buffer, offset, 0x6);
            m_flags = (FormulaFlags)BodyReadUInt16(buffer, offset, 0xE);
            m_formulaLength = BodyReadByte(buffer, offset, 0xF);
            int avail = Math.Max(0, length - 0x10);
            int n = Math.Min(m_formulaLength, avail);
            m_formulaBytes = new byte[n];
            if (n > 0) {
                Buffer.BlockCopy(buffer, offset + 0x10, m_formulaBytes, 0, n);
            }
            m_formula = Exceltk.Extension.DefaultEncoding().GetString(m_formulaBytes, 0, m_formulaBytes.Length);
        }

        protected override int GetRequiredEncodeBodyBufferLength() {
            return 0x10 + m_formulaBytes.Length;
        }

        protected override void EncodeBody(byte[] buffer, int offset, int capacity, out int written) {
            int need = GetRequiredEncodeBodyBufferLength();
            if (capacity < need) {
                throw new ArgumentException(Errors.ErrorBIFFBufferSize);
            }
            WriteUInt16(buffer, offset + 0x0, RowIndex);
            WriteUInt16(buffer, offset + 0x2, ColumnIndex);
            WriteUInt16(buffer, offset + 0x4, XFormat);
            WriteInt64(buffer, offset + 0x6, m_rawVal);
            WriteUInt16(buffer, offset + 0xE, (ushort)m_flags);
            buffer[offset + 0xF] = m_formulaLength;
            if (m_formulaBytes.Length > 0) {
                Buffer.BlockCopy(m_formulaBytes, 0, buffer, offset + 0x10, m_formulaBytes.Length);
            }
            written = need;
        }

        public Encoding UseEncoding {
            get {
                return m_UseEncoding;
            }
            set {
                m_UseEncoding = value;
            }
        }

        public FormulaFlags Flags {
            get {
                return m_flags;
            }
        }

        public byte FormulaLength {
            get {
                return m_formulaLength;
            }
        }

        /// <summary>
        /// Type-dependent formula result. String results look ahead to the next STRING package
        /// on the shared buffer (multi-record assembly).
        /// </summary>
        public object Value {
            get {
                long val = m_rawVal;
                if (((ulong)val & 0xFFFF000000000000) == 0xFFFF000000000000) {
                    var type = (byte)(val & 0xFF);
                    var code = (byte)((val >> 16) & 0xFF);
                    switch (type) {
                        case 0: // String
                            BinaryPackage rec = GetRecord(m_bytes, (uint)(Offset + Size), reader);
                            XlsBiffFormulaString str;
                            if (rec.ID == BIFFRECORDTYPE.SHRFMLA) {
                                str = GetRecord(m_bytes, (uint)(Offset + Size + rec.Size), reader) as
                                    XlsBiffFormulaString;
                            } else {
                                str = rec as XlsBiffFormulaString;
                            }

                            if (str == null) {
                                return string.Empty;
                            }
                            str.UseEncoding = m_UseEncoding;
                            return str.Value;
                        case 1: // Boolean
                            return (code != 0);
                        case 2: // Error
                            return (FORMULAERROR)code;
                        default:
                            return null;
                    }
                }
                return val.Int64BitsToDouble();
            }
        }

        public string Formula {
            get {
                return m_formula;
            }
        }
    }
}
