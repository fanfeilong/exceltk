using System;

namespace Exceltk.Reader.Package {
    /// <summary>
    /// Opened workbook container (ZIP package, OLE compound file, or plain file).
    /// Parsers read workbook content through this abstraction.
    /// </summary>
    public interface IExcelPackage : IDisposable {
        bool IsValid {
            get;
        }

        string ExceptionMessage {
            get;
        }
    }
}
