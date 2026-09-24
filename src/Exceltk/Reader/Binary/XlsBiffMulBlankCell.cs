using System;

namespace Exceltk.Reader.Binary {
    /// <summary>
    /// BIFF Package: MULBLANK — row, firstCol, rgixfe[ ], lastCol.
    /// </summary>
    internal class XlsBiffMulBlankCell : XlsBiffBlankCell {
        private ushort m_lastColumnIndex;
        private ushort[] m_xfByColumn = Array.Empty<ushort>();

        internal XlsBiffMulBlankCell(ExcelBinaryReader reader)
            : base(reader) {
        }

        internal XlsBiffMulBlankCell(byte[] bytes, uint offset, ExcelBinaryReader reader)
            : base(bytes, offset, reader) {
        }

        protected override void DecodeBody(byte[] buffer, int offset, int length) {
            // row + firstCol (+ first xf at 0x4 for BlankCell mid-base)
            base.DecodeBody(buffer, offset, length);
            if (length < 8) {
                m_lastColumnIndex = ColumnIndex;
                m_xfByColumn = Array.Empty<ushort>();
                return;
            }

            m_lastColumnIndex = BodyReadUInt16(buffer, offset, length - 2);
            int count = m_lastColumnIndex - ColumnIndex + 1;
            if (count < 0) {
                count = 0;
            }

            // Office MULBLANK: 2-byte ixfe per column starting at body+4.
            m_xfByColumn = new ushort[count];
            for (int i = 0; i < count; i++) {
                int ofs = 4 + 2 * i;
                if (ofs + 2 > length - 2) {
                    m_xfByColumn[i] = 0;
                } else {
                    m_xfByColumn[i] = BodyReadUInt16(buffer, offset, ofs);
                }
            }
        }

        protected override int GetRequiredEncodeBodyBufferLength() {
            // row(2)+colFirst(2)+xf[n]*2+colLast(2)
            return 4 + m_xfByColumn.Length * 2 + 2;
        }

        protected override void EncodeBody(byte[] buffer, int offset, int capacity, out int written) {
            int need = GetRequiredEncodeBodyBufferLength();
            if (capacity < need) {
                throw new ArgumentException(Errors.ErrorBIFFBufferSize);
            }
            WriteUInt16(buffer, offset + 0x0, RowIndex);
            WriteUInt16(buffer, offset + 0x2, ColumnIndex);
            for (int i = 0; i < m_xfByColumn.Length; i++) {
                WriteUInt16(buffer, offset + 4 + 2 * i, m_xfByColumn[i]);
            }
            WriteUInt16(buffer, offset + 4 + m_xfByColumn.Length * 2, m_lastColumnIndex);
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
    }
}
