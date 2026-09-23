using System;
using Exceltk.Reader.Package;

namespace Exceltk.Reader.Parser {
    /// <summary>
    /// Parses a binary (.xls) package via <see cref="ExcelBinaryReader"/>.
    /// </summary>
    public sealed class BinaryPackageParser : IExcelPackageParser {
        public DataSet Parse(IExcelPackage package) {
            var streamPackage = package as IStreamPackage;
            if (streamPackage == null) {
                throw new ArgumentException("BinaryPackageParser requires IStreamPackage", "package");
            }
            if (!streamPackage.IsValid) {
                return null;
            }

            var reader = new ExcelBinaryReader();
            try {
                reader.Open(streamPackage.Content);
                return reader.AsDataSet();
            } finally {
                reader.Close();
            }
        }
    }
}
