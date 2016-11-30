namespace Exceltk {
    public static class Config {
        // all mode
        public static int DecimalPrecision{
            get;
            set;
        }

        // all mode
        public static string DecimalFormat{
            get{
                return string.Format("{{0:N{0}}}", DecimalPrecision);
            }
        }
        
        // for md mode
        public static bool BodyHead {
            get;
            set;
        }

        // for tex mode
        public static bool SplitNumber {
            get;
            set;
        }

        // for tex mode
        public static bool SplitTable {
            get;
            set;
        }

        public static int SplitTableRow {
            get;
            set;
        }
    }
}
