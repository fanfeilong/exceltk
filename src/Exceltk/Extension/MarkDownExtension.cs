using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace ExcelToolKit{
    public class MarkDownTable {
        public string Name {
            get;
            set;
        }
        public string Value {
            get;
            set;
        }
    }

    public static class MarkDownExtension {
        public static MarkDownTable ToMd(this string xls, string sheet) {
            FileStream stream=File.Open(xls, FileMode.Open, FileAccess.Read);
            IExcelDataReader excelReader=null;
            if (Path.GetExtension(xls)==".xls") {
                excelReader=ExcelReaderFactory.CreateBinaryReader(stream);
            } else if (Path.GetExtension(xls)==".xlsx") {
                excelReader=ExcelReaderFactory.CreateOpenXmlReader(stream);
            } else {
                throw new ArgumentException("Not Support Format: ");
            }
            DataSet dataSet=excelReader.AsDataSet();
            DataTable dataTable=dataSet.Tables[sheet];

            var table=new MarkDownTable {
                    Name=dataTable.TableName,
                    Value=dataTable.ToMd()
            };

            excelReader.Close();

            return table;
        }
        public static IEnumerable<MarkDownTable> ToMd(this string xls) {
            FileStream stream=File.Open(xls, FileMode.Open, FileAccess.Read);
            IExcelDataReader excelReader=null;
            if (Path.GetExtension(xls)==".xls") {
                excelReader=ExcelReaderFactory.CreateBinaryReader(stream);
            } else if (Path.GetExtension(xls)==".xlsx") {
                excelReader=ExcelReaderFactory.CreateOpenXmlReader(stream);
            } else {
                throw new ArgumentException("Not Support Format: ");
            }
            DataSet dataSet=excelReader.AsDataSet();

            foreach (DataTable dataTable in dataSet.Tables) {
                var table=new MarkDownTable {
                        Name=dataTable.TableName,
                        Value=dataTable.ToMd()
                };

                yield return table;
            }

            excelReader.Close();
        }

        public static string ToMd(this DataTable table, bool insertHeader=true) {
            table.Shrink();
            //table.RemoveColumnsByRow(0, string.IsNullOrEmpty);
            var sb=new StringBuilder();

            int i=0;
            foreach (DataRow row in table.Rows) {
                sb.Append("|");
                foreach (object cell in row.ItemArray) {
                    string value=GetCellValue(cell);
                    sb.Append(value).Append("|");
                }

                sb.Append("\r\n");
                if (i==0&&insertHeader) {
                    sb.Append("|");
                    foreach (DataColumn col in table.Columns) {
                        sb.Append(":--|");
                    }
                    sb.Append("\r\n");
                }
                i++;
            }
            return sb.ToString();
        }
        private static string GetCellValue(object cell) {
            if (cell==null) {
                return "";
            }
            string value;
            var xlsCell=cell as XlsCell;
            if (xlsCell!=null) {
                value=xlsCell.MarkDownText;
            } else {
                value=cell.ToString();
            }

            // Decimal precision
            if (Config.DecimalPrecision>0) {
                if (Regex.IsMatch(value, @"^(-?[0-9]{1,}[.][0-9]*)$")) {
                    var old=value;
                    value=string.Format(Config.DecimalFormat, Double.Parse(value));
                }
            }

            return value;
        }
    }
}