using System;
using System.IO;
using Exceltk.Reader.Package;

namespace Exceltk.Reader.Parser {
    /// <summary>
    /// Creates the matching package + parser pair for a workbook path.
    /// </summary>
    public static class ExcelPackageParserFactory {
        public static IExcelPackage OpenPackage(string path) {
            string ext = Path.GetExtension(path);
            if (string.Equals(ext, ".csv", StringComparison.OrdinalIgnoreCase)) {
                return new PathPackage(path);
            }
            if (string.Equals(ext, ".xls", StringComparison.OrdinalIgnoreCase)) {
                Stream stream = File.Open(path, FileMode.Open, FileAccess.Read);
                return new StreamPackage(stream, ownsStream: true);
            }
            if (string.Equals(ext, ".xlsx", StringComparison.OrdinalIgnoreCase)) {
                Stream stream = File.Open(path, FileMode.Open, FileAccess.Read);
                var package = new ZipWorker();
                package.Extract(stream);
                return package;
            }
            throw new ArgumentException("Not Support Format: " + ext);
        }

        public static IExcelPackageParser CreateParser(string path) {
            string ext = Path.GetExtension(path);
            if (string.Equals(ext, ".csv", StringComparison.OrdinalIgnoreCase)) {
                return new CsvPackageParser();
            }
            if (string.Equals(ext, ".xls", StringComparison.OrdinalIgnoreCase)) {
                return new BinaryPackageParser();
            }
            if (string.Equals(ext, ".xlsx", StringComparison.OrdinalIgnoreCase)) {
                return new OpenXmlPackageParser();
            }
            throw new ArgumentException("Not Support Format: " + ext);
        }

        public static DataSet Load(string path) {
            using (IExcelPackage package = OpenPackage(path)) {
                IExcelPackageParser parser = CreateParser(path);
                return parser.Parse(package);
            }
        }
    }
}
