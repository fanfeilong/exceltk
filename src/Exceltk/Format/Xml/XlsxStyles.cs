using System.Collections.Generic;

namespace ExcelToolKit.Format.Xml {
    internal class XlsxStyles {
        public XlsxStyles() {
            CellXfs=new List<XlsxXf>();
            NumFmts=new List<XlsxNumFmt>();
        }

        public List<XlsxXf> CellXfs {
            get;
            set;
        }

        public List<XlsxNumFmt> NumFmts {
            get;
            set;
        }
    }
}