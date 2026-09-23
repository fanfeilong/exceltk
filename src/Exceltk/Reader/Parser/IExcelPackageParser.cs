using Exceltk.Reader.Package;

namespace Exceltk.Reader.Parser {
    /// <summary>
    /// Parses an <see cref="IExcelPackage"/> into the shared <see cref="DataSet"/> model.
    /// </summary>
    public interface IExcelPackageParser {
        DataSet Parse(IExcelPackage package);
    }
}
