namespace Exceltk.Reader.Xml {
    internal class XlsxNumFmt {
        public const string N_numFmt="numFmt";
        public const string A_numFmtId="numFmtId";
        public const string A_formatCode="formatCode";

        public XlsxNumFmt(int id, string formatCode) {
            Id=id;
            FormatCode=formatCode;
        }

        public int Id {
            get;
            set;
        }

        public string FormatCode {
            get;
            set;
        }
    }
}