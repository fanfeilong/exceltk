using System;
using Exceltk.Reader.Package;

namespace Exceltk.Reader.Parser {
    /// <summary>
    /// Parses a CSV path package via <see cref="CsvReader"/>.
    /// </summary>
    public sealed class CsvPackageParser : IExcelPackageParser {
        public DataSet Parse(IExcelPackage package) {
            var pathPackage = package as IPathPackage;
            if (pathPackage == null) {
                throw new ArgumentException("CsvPackageParser requires IPathPackage", "package");
            }
            if (!pathPackage.IsValid) {
                return null;
            }
            return CsvReader.AsDataSet(pathPackage.Path);
        }
    }
}
