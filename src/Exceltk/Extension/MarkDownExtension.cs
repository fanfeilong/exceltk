using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

using Exceltk.Reader;

namespace Exceltk{
    public static class MarkDownExtension {
        public static SimpleTable ToMd(this string xls, string sheet) {
            DataSet dataSet=WorkbookLoader.Load(xls);
            DataTable dataTable=ResolveTable(dataSet, sheet);

            return new SimpleTable {
                    Name=dataTable.TableName,
                    Value=dataTable.ToMd(dataSet)
            };
        }
        public static IEnumerable<SimpleTable> ToMd(this string xls) {
            DataSet dataSet=WorkbookLoader.Load(xls);

            foreach (DataTable dataTable in dataSet.Tables) {
                yield return new SimpleTable {
                        Name=dataTable.TableName,
                        Value=dataTable.ToMd(dataSet)
                };
            }
        }

        private static DataTable ResolveTable(DataSet dataSet, string sheet) {
            if (dataSet.Tables.ContainsTable(sheet)) {
                return dataSet.Tables[sheet];
            }
            // CSV has a single unnamed table; allow -sheet to be omitted or ignored when only one table exists.
            if (dataSet.Tables.Count==1 && (string.IsNullOrEmpty(sheet) || string.IsNullOrEmpty(dataSet.Tables[0].TableName))) {
                return dataSet.Tables[0];
            }
            throw new ArgumentException("Sheet not found: "+sheet);
        }

        public static string ToMd(this DataTable table, DataSet dataSet, bool insertHeader=true) {
            table.Shrink();
            //table.RemoveColumnsByRow(0, string.IsNullOrEmpty);
            var sb=new StringBuilder();

            int i=0;
            foreach (DataRow row in table.Rows) {
                if (Config.BodyHead) {
                    if (i == 0 && insertHeader) {
                        sb.Append("|");
                        foreach (DataColumn col in table.Columns) {
                            sb.Append("|");
                        }
                        sb.Append(Environment.NewLine);

                        sb.Append("|");
                        foreach (DataColumn col in table.Columns) {
                            sb.Append(Config.TableAliginFormat).Append("|");
                        }
                        sb.Append(Environment.NewLine);
                    }
                }

                sb.Append("|");
                foreach (object cell in row.ItemArray) {
                    string value=GetCellValue(dataSet, cell);
                    if (i == 0 && Config.BodyHead) {
                        sb.AppendFormat("**{0}**",value).Append("|");
                    } else {
                        sb.Append(value).Append("|");
                    }
                }

                sb.Append(Environment.NewLine);

                if (!Config.BodyHead) {
                    if (i == 0 && insertHeader) {
                        sb.Append("|");
                        foreach (DataColumn col in table.Columns) {
                            sb.Append(Config.TableAliginFormat).Append("|");
                        }
                        sb.Append(Environment.NewLine);
                    }
                }

                i++;
            }
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

            //Console.WriteLine(value);

            // Decimal precision
            if (Config.HasDecimalPrecision) {
                if (Regex.IsMatch(value, @"^(-?[0-9]{1,}[.][0-9]*)$")) {
                    var old=value;
                    if(Config.DecimalPrecision>0){
                        value = string.Format(Config.DecimalFormat, Double.Parse(value));
                    }else{
                        value = ((int)Double.Parse(value)).ToString();
                    }
                }
            }

            value = Regex.Replace(value, @"\r\n?|\n", "<br/>");
            value = value.Replace("|", "\\|");

            return value;
        }
    }
}