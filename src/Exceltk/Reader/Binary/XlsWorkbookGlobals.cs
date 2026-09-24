using System;
using System.Collections.Generic;
using System.Linq;

namespace Exceltk.Reader.Binary {
    /// <summary>
    /// Represents Globals section of workbook
    /// </summary>
    internal class XlsWorkbookGlobals {
        private readonly List<BinaryPackage> m_ExtendedFormats=new List<BinaryPackage>();
        private readonly List<BinaryPackage> m_Fonts=new List<BinaryPackage>();

        private readonly Dictionary<ushort, XlsBiffFormatString> m_Formats=
            new Dictionary<ushort, XlsBiffFormatString>();

        private readonly List<XlsBiffHyperLink> m_HyperLinkTable;

        private readonly List<XlsBiffBoundSheet> m_Sheets=new List<XlsBiffBoundSheet>();
        private readonly List<BinaryPackage> m_Styles=new List<BinaryPackage>();

        public XlsWorkbookGlobals() {
            m_HyperLinkTable=new List<XlsBiffHyperLink>();
        }

        public XlsBiffInterfaceHdr InterfaceHdr {
            get;
            set;
        }

        public BinaryPackage MMS {
            get;
            set;
        }

        public BinaryPackage WriteAccess {
            get;
            set;
        }

        public XlsBiffSimpleValueRecord CodePage {
            get;
            set;
        }

        public BinaryPackage DSF {
            get;
            set;
        }

        public BinaryPackage Country {
            get;
            set;
        }

        public XlsBiffSimpleValueRecord Backup {
            get;
            set;
        }

        public List<BinaryPackage> Fonts {
            get {
                return m_Fonts;
            }
        }

        public Dictionary<ushort, XlsBiffFormatString> Formats {
            get {
                return m_Formats;
            }
        }


        public List<BinaryPackage> ExtendedFormats {
            get {
                return m_ExtendedFormats;
            }
        }

        public List<BinaryPackage> Styles {
            get {
                return m_Styles;
            }
        }

        public List<XlsBiffBoundSheet> Sheets {
            get {
                return m_Sheets;
            }
        }

        /// <summary>
        /// Shared String Table of workbook
        /// </summary>
        public XlsBiffSST SST {
            get;
            set;
        }

        public BinaryPackage ExtSST {
            get;
            set;
        }

        public void AddHyperLink(XlsBiffHyperLink hyperLink) {
            if(hyperLink.Url!="_blank"){
                Console.WriteLine(hyperLink.Url);
            }
            m_HyperLinkTable.Add(hyperLink);
        }

        public XlsBiffHyperLink GetHyperLink(int row, int column){
            return m_HyperLinkTable.FirstOrDefault(h => h.CellRangeAddress.Contain(row, column));
        }
    }
}