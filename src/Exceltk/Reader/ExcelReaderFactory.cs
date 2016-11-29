using System.IO;

namespace Exceltk.Reader {
    public static class ExcelReaderFactory {
        public static IExcelDataReader CreateReader(Stream fileStream, ExcelFileType excelFileType) {
            IExcelDataReader reader=null;

            switch (excelFileType) {
                case ExcelFileType.Binary:
                    reader=new ExcelBinaryReader();
                    reader.Open(fileStream);
                    break;
                case ExcelFileType.OpenXml:
                    reader=new ExcelOpenXmlReader();
                    reader.Open(fileStream);
                    break;
                default:
                    break;
            }

            return reader;
        }

        public static IExcelDataReader CreateBinaryReader(Stream fileStream) {
            IExcelDataReader reader=new ExcelBinaryReader();
            reader.Open(fileStream);

            return reader;
        }

        public static IExcelDataReader CreateBinaryReader(Stream fileStream, ReadOption option) {
            IExcelDataReader reader=new ExcelBinaryReader(option);
            reader.Open(fileStream);

            return reader;
        }

        public static IExcelDataReader CreateOpenXmlReader(Stream fileStream) {
            IExcelDataReader reader=new ExcelOpenXmlReader();
            reader.Open(fileStream);

            return reader;
        }
    }
}