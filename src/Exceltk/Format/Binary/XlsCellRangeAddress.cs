using System;

namespace ExcelToolKit.Format.Binary {
    internal class XlsCellRangeAddress {
        protected byte[] m_bytes;
        protected int m_offset;

        public XlsCellRangeAddress(byte[] bytes, int offset) {
            m_bytes=bytes;
            m_offset=offset;
        }

        public int FirstRow {
            get {
                return BitConverter.ToInt16(m_bytes, m_offset+0);
            }
        }

        public int LastRow {
            get {
                return BitConverter.ToInt16(m_bytes, m_offset+2);
            }
        }

        public int FirstColumin {
            get {
                return BitConverter.ToInt16(m_bytes, m_offset+4);
            }
        }

        public int LastColumin {
            get {
                return BitConverter.ToInt16(m_bytes, m_offset+6);
            }
        }

        public override string ToString() {
            return string.Format("[<{0},{1}>,<{2},{3}>],", FirstRow, FirstColumin, LastRow, LastColumin);
        }

        public bool Contain(int row, int column) {
            return FirstRow<=row&&row<=LastRow&&FirstColumin<=column&&column<=LastColumin;
        }
    }
}