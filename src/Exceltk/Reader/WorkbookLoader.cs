using System;
using System.IO;

namespace Exceltk.Reader {
    /// <summary>
    /// Loads a workbook-like DataSet from .xls, .xlsx, or .csv.
    /// Binary/XML formats are read via streaming package parsers internally.
    /// </summary>
    public static class WorkbookLoader {
        public static DataSet Load(string path) {
            string ext=Path.GetExtension(path);
            if (string.Equals(ext, ".csv", StringComparison.OrdinalIgnoreCase)) {
                return CsvReader.AsDataSet(path);
            }

            FileStream stream=File.Open(path, FileMode.Open, FileAccess.Read);
            IExcelDataReader excelReader=null;
            try {
                if (string.Equals(ext, ".xls", StringComparison.OrdinalIgnoreCase)) {
                    excelReader=ExcelReaderFactory.CreateBinaryReader(stream);
                } else if (string.Equals(ext, ".xlsx", StringComparison.OrdinalIgnoreCase)) {
                    excelReader=ExcelReaderFactory.CreateOpenXmlReader(stream);
                } else {
                    stream.Close();
                    throw new ArgumentException("Not Support Format: "+ext);
                }
                DataSet dataSet=excelReader.AsDataSet();
                return dataSet;
            } finally {
                if (excelReader!=null) {
                    excelReader.Close();
                } else {
                    stream.Close();
                }
            }
        }

        public static bool IsSupportedExtension(string path) {
            string ext=Path.GetExtension(path);
            return string.Equals(ext, ".xls", StringComparison.OrdinalIgnoreCase)
                || string.Equals(ext, ".xlsx", StringComparison.OrdinalIgnoreCase)
                || string.Equals(ext, ".csv", StringComparison.OrdinalIgnoreCase);
        }
    }
}
