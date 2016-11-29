using System.IO;

namespace Exceltk.Reader {
    public interface IExcelDataReader {
        void Open(Stream fileStream);
        DataSet AsDataSet();
        void Close();
    }
}