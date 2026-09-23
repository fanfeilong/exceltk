using System;
using Exceltk.Reader.Package;

namespace Exceltk.Reader.Parser {
    /// <summary>
    /// Parses an OpenXML (.xlsx) package via <see cref="ExcelOpenXmlReader"/>.
    /// </summary>
    public sealed class OpenXmlPackageParser : IExcelPackageParser {
        public DataSet Parse(IExcelPackage package) {
            var openXml = package as IOpenXmlPackage;
            if (openXml == null) {
                throw new ArgumentException("OpenXmlPackageParser requires IOpenXmlPackage", "package");
            }
            if (!openXml.IsValid) {
                return null;
            }

            var reader = new ExcelOpenXmlReader();
            try {
                reader.Open(openXml);
                return reader.AsDataSet();
            } finally {
                reader.Close();
            }
        }
    }
}
