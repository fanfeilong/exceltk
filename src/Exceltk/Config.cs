namespace Exceltk {
    public static class Config {
        public static int DecimalPrecision{
            get;
            set;
        }
        public static string DecimalFormat{
            get{
                return string.Format("{{0:N{0}}}", DecimalPrecision);
            }
        }
        public static bool BodyHead {
            get;
            set;
        }
    }
}
