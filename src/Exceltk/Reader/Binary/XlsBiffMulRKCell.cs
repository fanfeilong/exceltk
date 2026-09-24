using System;

namespace Exceltk.Reader.Binary {
    /// <summary>
    /// BIFF Package: MULRK — row, firstCol, (xf+rk)*n, lastCol.
    /// </summary>
    internal class XlsBiffMulRKCell : XlsBiffBlankCell {
        private ushort m_lastColumnIndex;
        private ushort[] m_xfByColumn = Array.Empty<ushort>();
        private uint[] m_rkByColumn = Array.Empty<uint>();
        private double[] m_valueByColumn = Array.Empty<double>();

        internal XlsBiffMulRKCell(ExcelBinaryReader reader)
            : base(reader) {
        }

        internal XlsBiffMulRKCell(byte[] bytes, uint offset, ExcelBinaryReader reader)
            : base(bytes, offset, reader) {
        }

        protected override void DecodeBody(byte[] buffer, int offset, int length) {
            base.DecodeBody(buffer, offset, length);
            if (length < 8) {
                m_lastColumnIndex = ColumnIndex;
                m_xfByColumn = Array.Empty<ushort>();
                m_rkByColumn = Array.Empty<uint>();
                m_valueByColumn = Array.Empty<double>();
                return;
            }

            m_lastColumnIndex = BodyReadUInt16(buffer, offset, length - 2);
            int count = m_lastColumnIndex - ColumnIndex + 1;
            if (count < 0) {
                count = 0;
            }

            m_xfByColumn = new ushort[count];
            m_rkByColumn = new uint[count];
            m_valueByColumn = new double[count];
            for (int i = 0; i < count; i++) {
                int xfOfs = 4 + 6 * i;
                int valOfs = 6 + 6 * i;
                if (xfOfs + 2 > length - 2) {
                    m_xfByColumn[i] = 0;
                    m_rkByColumn[i] = 0;
                    m_valueByColumn[i] = 0;
                } else {
                    m_xfByColumn[i] = BodyReadUInt16(buffer, offset, xfOfs);
                    m_rkByColumn[i] = valOfs + 4 <= length - 2
                        ? BodyReadUInt32(buffer, offset, valOfs)
                        : 0u;
                    m_valueByColumn[i] = XlsBiffRKCell.NumFromRK(m_rkByColumn[i]);
                }
            }
        }

        protected override int GetRequiredEncodeBodyBufferLength() {
            // row(2)+colFirst(2)+(xf+rk)*n + colLast(2)
            return 4 + m_rkByColumn.Length * 6 + 2;
        }

        protected override void EncodeBody(byte[] buffer, int offset, int capacity, out int written) {
            int need = GetRequiredEncodeBodyBufferLength();
            if (capacity < need) {
                throw new ArgumentException(Errors.ErrorBIFFBufferSize);
            }
            WriteUInt16(buffer, offset + 0x0, RowIndex);
            WriteUInt16(buffer, offset + 0x2, ColumnIndex);
            for (int i = 0; i < m_rkByColumn.Length; i++) {
                WriteUInt16(buffer, offset + 4 + 6 * i, m_xfByColumn[i]);
                WriteUInt32(buffer, offset + 6 + 6 * i, m_rkByColumn[i]);
            }
            WriteUInt16(buffer, offset + 4 + m_rkByColumn.Length * 6, m_lastColumnIndex);
            written = need;
        }

        public ushort LastColumnIndex {
            get {
                return m_lastColumnIndex;
            }
        }

        public ushort GetXF(ushort ColumnIdx) {
            int i = ColumnIdx - ColumnIndex;
            if (i < 0 || i >= m_xfByColumn.Length) {
                return 0;
            }
            return m_xfByColumn[i];
        }

        public double GetValue(ushort ColumnIdx) {
            int i = ColumnIdx - ColumnIndex;
            if (i < 0 || i >= m_valueByColumn.Length) {
                return 0;
            }
            return m_valueByColumn[i];
        }
    }
}
