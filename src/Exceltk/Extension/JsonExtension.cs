using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

using Exceltk.Reader;

namespace Exceltk
{
    public static class JsonExtension{
        public static SimpleTable ToJson(this string xls, string sheet) {
            DataSet dataSet=WorkbookLoader.Load(xls);
            DataTable dataTable=dataSet.Tables.ContainsTable(sheet)
                ? dataSet.Tables[sheet]
                : (dataSet.Tables.Count==1 ? dataSet.Tables[0] : dataSet.Tables[sheet]);

            return new SimpleTable {
                    Name=dataTable.TableName,
                    Value=dataTable.ToJson(dataSet)
            };
        }

        public static IEnumerable<SimpleTable> ToJson(this string xls) {
            DataSet dataSet=WorkbookLoader.Load(xls);

            foreach (DataTable dataTable in dataSet.Tables) {
                yield return new SimpleTable {
                        Name=dataTable.TableName,
                        Value=dataTable.ToJson(dataSet)
                };
            }
        }

        public static string ToJson(this DataTable table, DataSet dataSet, bool insertHeader=true) {
            table.Shrink();
            //table.RemoveColumnsByRow(0, string.IsNullOrEmpty);
            var sb=new StringBuilder();

            int i=0;
            sb.AppendLine("{");
            sb.AppendFormat("\t'name':'{0}',\n",table.TableName);
            sb.AppendFormat("\t'rows':[\n");
            var columns = new Dictionary<int,string>();
            foreach (DataRow row in table.Rows) {

                if(i==0){
                    int j=0;
                    foreach (object cell in row.ItemArray) {
                        string value=GetCellValue(dataSet, cell);
                        columns[j] = value;
                        j++;
                    }
                }else{
                    sb.Append("\t\t{\n");
                    int j=0;
                    foreach (object cell in row.ItemArray) {
                        string value=GetCellValue(dataSet, cell);
                        sb.AppendFormat("\t\t\t'{0}':'{1}'",columns[j],value);
                        if(j<row.ItemArray.Length-1){
                            sb.Append(",\n");
                        }
                        j++;
                    }
                    sb.Append("\n\t\t}");
                    if(i<table.Rows.Count-1){
                        sb.Append(",");
                    }
                    sb.Append("\n");
                }
                i++;
            }
            sb.AppendFormat("\t]\n");
            sb.AppendLine("}");
            return sb.ToString();
        }

        private static string GetCellValue(DataSet dataSet, object cell) {
            if (cell==null) {
                return "";
            }
            string value;
            var xlsCell=cell as XlsCell;
            if (xlsCell!=null) {
                value=xlsCell.GetMarkDownText(dataSet);
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

            value = value.Replace("\r","\\r");
            value = value.Replace("\n","\\n");
            value = value.Replace("\t","\\t");

            return value;
        }
    }
}