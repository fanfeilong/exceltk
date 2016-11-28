using System.IO;

namespace ExcelToolKit {
    public interface IExcelDataReader {
        void Open(Stream fileStream);
        DataSet AsDataSet();
        void Close();
    }
}