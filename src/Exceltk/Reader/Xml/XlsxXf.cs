namespace Exceltk.Reader.Xml {
    internal class XlsxXf {
        public const string N_xf="xf";
        public const string A_numFmtId="numFmtId";
        public const string A_xfId="xfId";
        public const string A_applyNumberFormat="applyNumberFormat";

        public XlsxXf(int id, int numFmtId, string applyNumberFormat) {
            Id=id;
            NumFmtId=numFmtId;
            ApplyNumberFormat=null!=applyNumberFormat&&applyNumberFormat=="1";
        }

        public int Id {
            get;
            set;
        }

        public int NumFmtId {
            get;
            set;
        }

        public bool ApplyNumberFormat {
            get;
            set;
        }
    }
}