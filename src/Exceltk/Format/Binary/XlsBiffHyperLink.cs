using System;
using System.Text;

namespace ExcelToolKit.Format.Binary {
    internal class XlsBiffHyperLink : XlsBiffRecord {
        private XlsCellRangeAddress m_cellRangeAddress;

        internal XlsBiffHyperLink(byte[] bytes, uint offset, ExcelBinaryReader reader)
            : base(bytes, offset, reader) {
        }

        public UInt32 Flags {
            get {
                return BitConverter.ToUInt32(m_bytes, 28);
            }
        }

        public bool HasUrl {
            get {
                return (Flags&0x00000001)==1;
            }
        }

        public bool IsRelative {
            get {
                return (Flags&0x00000002)>>1==1;
            }
        }

        public bool HasDescription {
            get {
                uint bit=(Flags&0x00000014);
                return bit>>2==1&&bit>>4==1;
            }
        }

        public bool HasTextMark {
            get {
                return (Flags&0x00000008)>>3==1;
            }
        }

        public bool HasTatgetFrame {
            get {
                return (Flags&0x00000080)>>7==1;
            }
        }

        public bool IsUNC {
            get {
                return (Flags&0x00000100)>>8==1;
            }
        }

        public bool IsFileOrUrl {
            get {
                return !IsUNC;
            }
        }


        public string Url {
            get {
                int offset=32+m_readoffset;

                if (HasDescription) {
                    var descriptCount=(int)BitConverter.ToUInt32(m_bytes, offset);
                    offset+=4;
                    offset+=descriptCount;
                }

                if (HasTatgetFrame) {
                    var frameCount=(int)BitConverter.ToUInt32(m_bytes, offset);
                    offset+=4;
                    offset+=frameCount;
                }

                var urlSize=(int)BitConverter.ToUInt32(m_bytes, offset);
                //Console.WriteLine(urlSize);
                offset+=4;

                //byte[] bytes = new byte[ urlSize*2];

                string value=Encoding.Unicode.GetString(m_bytes, offset, 2*(urlSize-1));
                ////Console.WriteLine("{0}**",value.Replace("\t").TrimEnd());
                return value;
            }
        }

        public XlsCellRangeAddress CellRangeAddress {
            get {
                if (m_cellRangeAddress==null) {
                    m_cellRangeAddress=new XlsCellRangeAddress(m_bytes, m_readoffset);
                }
                return m_cellRangeAddress;
            }
        }
    }
}