using System.IO;

namespace ExcelToolKit {
    public static class ExcelReaderFactory {
        public static IExcelDataReader CreateReader(Stream fileStream, ExcelFileType excelFileType) {
            IExcelDataReader reader=null;

            switch (excelFileType) {
                case ExcelFileType.Binary:
                    reader=new ExcelBinaryReader();
                    reader.Initialize(fileStream);
                    break;
                case ExcelFileType.OpenXml:
                    reader=new ExcelOpenXmlReader();
                    reader.Initialize(fileStream);
                    break;
                default:
                    break;
            }

            return reader;
        }

        public static IExcelDataReader CreateBinaryReader(Stream fileStream) {
            IExcelDataReader reader=new ExcelBinaryReader();
            reader.Initialize(fileStream);

            return reader;
        }

        public static IExcelDataReader CreateBinaryReader(Stream fileStream, ReadOption option) {
            IExcelDataReader reader=new ExcelBinaryReader(option);
            reader.Initialize(fileStream);

            return reader;
        }

        public static IExcelDataReader CreateBinaryReader(Stream fileStream, bool convertOADate) {
            IExcelDataReader reader=CreateBinaryReader(fileStream);
            ((ExcelBinaryReader)reader).ConvertOaDate=convertOADate;

            return reader;
        }

        public static IExcelDataReader CreateBinaryReader(Stream fileStream, bool convertOADate, ReadOption readOption) {
            IExcelDataReader reader=CreateBinaryReader(fileStream, readOption);
            ((ExcelBinaryReader)reader).ConvertOaDate=convertOADate;

            return reader;
        }

        public static IExcelDataReader CreateOpenXmlReader(Stream fileStream) {
            IExcelDataReader reader=new ExcelOpenXmlReader();
            reader.Initialize(fileStream);

            return reader;
        }
    }
}