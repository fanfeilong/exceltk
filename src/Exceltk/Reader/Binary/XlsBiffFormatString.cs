using System.Text;

namespace Exceltk.Reader.Binary {
    /// <summary>
    /// Represents a string value of formula
    /// </summary>
    internal class XlsBiffFormatString : XlsBiffRecord {
        private Encoding m_UseEncoding=Exceltk.Extension.DefaultEncoding();
        private string m_value;

        internal XlsBiffFormatString(byte[] bytes, uint offset, ExcelBinaryReader reader)
            : base(bytes, offset, reader) {
        }


        /// <summary>
        /// Encoding used to deal with strings
        /// </summary>
        public Encoding UseEncoding {
            get {
                return m_UseEncoding;
            }
            set {
                m_UseEncoding=value;
            }
        }

        /// <summary>
        /// Length of the string
        /// </summary>
        public ushort Length {
            get {
                switch (ID) {
                    case BIFFRECORDTYPE.FORMAT_V23:
                        return base.ReadByte(0x0);
                    default:
                        return base.ReadUInt16(2);
                }
            }
        }

        /// <summary>
        /// String text
        /// </summary>
        public string Value {
            get {
                if (m_value==null) {
                    switch (ID) {
                        case BIFFRECORDTYPE.FORMAT_V23:
                            m_value=m_UseEncoding.GetString(m_bytes, m_readoffset+1, Length);
                            break;
                        case BIFFRECORDTYPE.FORMAT:
                            int offset=m_readoffset+5;
                            byte flags=ReadByte(3);
                            m_UseEncoding=(flags&0x01)==0x01?Encoding.Unicode:Exceltk.Extension.DefaultEncoding();
                            if ((flags&0x04)==0x01) // asian phonetic block size
                                offset+=4;
                            if ((flags&0x08)==0x01) // number of rtf blocks
                                offset+=2;
                            m_value=m_UseEncoding.IsSingleByte
                                          ?m_UseEncoding.GetString(m_bytes, offset, Length)
                                          :m_UseEncoding.GetString(m_bytes, offset, Length*2);

                            break;
                    }
                }
                return m_value;
            }
        }

        public ushort Index {
            get {
                switch (ID) {
                    case BIFFRECORDTYPE.FORMAT_V23:
                        return 0;
                    default:
                        return ReadUInt16(0);
                }
            }
        }
    }
}