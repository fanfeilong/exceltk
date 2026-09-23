using System;
using System.IO;
using Exceltk.Reader.Parser;

namespace Exceltk.Reader {
    /// <summary>
    /// Loads a workbook-like DataSet from .xls, .xlsx, or .csv via the
    /// package / package-parser pipeline.
    /// </summary>
    public static class WorkbookLoader {
        public static DataSet Load(string path) {
            return ExcelPackageParserFactory.Load(path);
        }

        public static bool IsSupportedExtension(string path) {
            string ext=Path.GetExtension(path);
            return string.Equals(ext, ".xls", StringComparison.OrdinalIgnoreCase)
                || string.Equals(ext, ".xlsx", StringComparison.OrdinalIgnoreCase)
                || string.Equals(ext, ".csv", StringComparison.OrdinalIgnoreCase);
        }
    }
}
