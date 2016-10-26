using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelToolKit {
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
    }
}
