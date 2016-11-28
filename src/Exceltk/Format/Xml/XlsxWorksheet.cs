namespace ExcelToolKit.Format.Xml {
    internal class XlsxWorksheet {
        public const string N_dimension="dimension";
        public const string N_worksheet="worksheet";
        public const string N_row="row";
        public const string N_col="col";
        public const string N_c="c"; //cell
        public const string N_v="v";
        public const string N_f="f"; //
        public const string N_t="t";
        public const string A_ref="ref";
        public const string A_r="r";
        public const string A_t="t";
        public const string A_s="s";
        public const string A_display="display";
        public const string A_rid="r:id";
        public const string N_sheetData="sheetData";
        public const string N_inlineStr="inlineStr";
        public const string N_hyperlinks="hyperlinks";
        public const string N_hyperlink="hyperlink";
        private readonly string _Name;
        private readonly int _id;

        private XlsxDimension _dimension;

        public XlsxWorksheet(string name, int id, string rid) {
            _Name=name;
            _id=id;
            RID=rid;
        }

        public bool IsEmpty {
            get;
            set;
        }

        public XlsxDimension Dimension {
            get {
                return _dimension;
            }
            set {
                _dimension=value;
            }
        }

        public int ColumnsCount {
            get {
                return IsEmpty?0:(_dimension==null?-1:_dimension.LastCol);
            }
        }

        public int RowsCount {
            get {
                return _dimension==null?-1:_dimension.LastRow-_dimension.FirstRow+1;
            }
        }

        public string Name {
            get {
                return _Name;
            }
        }

        public int Id {
            get {
                return _id;
            }
        }

        public string RID {
            get;
            set;
        }


        public string Path {
            get;
            set;
        }
    }
}