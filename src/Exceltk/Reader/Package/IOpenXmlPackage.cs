using System.IO;

namespace Exceltk.Reader.Package {
    /// <summary>
    /// OpenXML (.xlsx) package: ZIP parts for workbook, sheets, shared strings, and styles.
    /// </summary>
    public interface IOpenXmlPackage : IExcelPackage {
        Stream GetSharedStringsStream();
        Stream GetStylesStream();
        Stream GetWorkbookStream();
        Stream GetWorkbookRelsStream();
        Stream GetWorksheetStream(string sheetPath);
        Stream GetWorksheetRelsStream(string sheetPath);
    }
}
