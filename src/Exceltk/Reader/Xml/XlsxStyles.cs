using System.Collections.Generic;

namespace Exceltk.Reader.Xml {
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