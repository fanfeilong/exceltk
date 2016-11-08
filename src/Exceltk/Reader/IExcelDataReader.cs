using System.IO;

namespace ExcelToolKit {
    public interface IExcelDataReader {
        bool IsValid {get;}
        string ExceptionMessage {get;}
        string Name {get;}
        int ResultsCount {get;}
        bool IsFirstRowAsColumnNames {get;set;}

        void Initialize(Stream fileStream);
        DataSet AsDataSet();
        DataSet AsDataSet(bool convertOADateTime);
        void Close();
    }
}