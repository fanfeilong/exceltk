using System.IO;

namespace Exceltk.Reader.Package {
    /// <summary>
    /// Package that exposes a single content stream (binary .xls or similar).
    /// </summary>
    public interface IStreamPackage : IExcelPackage {
        Stream Content {
            get;
        }
    }
}
