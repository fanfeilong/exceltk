using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

using Exceltk.Reader;

namespace Exceltk
{
    public static class TexExtension{
        public static SimpleTable ToTex(this string xls, string sheet) {
            DataSet dataSet=WorkbookLoader.Load(xls);
            DataTable dataTable=dataSet.Tables.ContainsTable(sheet)
                ? dataSet.Tables[sheet]
                : (dataSet.Tables.Count==1 ? dataSet.Tables[0] : dataSet.Tables[sheet]);

            return new SimpleTable {
                    Name=dataTable.TableName,
                    Value=dataTable.ToTex(dataSet)
            };
        }

        public static IEnumerable<SimpleTable> ToTex(this string xls) {
            DataSet dataSet=WorkbookLoader.Load(xls);

            foreach (DataTable dataTable in dataSet.Tables) {
                yield return new SimpleTable {
                        Name=dataTable.TableName,
                        Value=dataTable.ToTex(dataSet)
                };
            }
        }

        public static string ToTex(this DataTable table, DataSet dataSet, bool insertHeader = true) {
            table.Shrink();
            //table.RemoveColumnsByRow(0, string.IsNullOrEmpty);
            var sb = new StringBuilder();

            int i = 0;
            sb.AppendLine(@"\newlength\q");
            sb.AppendLine(@"\setlength\q{\dimexpr(0.5\textwidth + 0.5\tabcolsep)/" + table.Columns.Count + "}");

            sb.AppendLine("\\begin{center}");

            var islongTable = false;
            foreach (DataRow row in table.Rows) {

                if (row.ItemArray.Length == 0) {
                    continue;
                }

                if (Config.SplitTable) {
                    if (i == Config.SplitTableRow) {
                        i = 0;
                        sb.Append("\t\\end{tabular}\n");
                    }
                }


                if(i==0){

                    if (row.ItemArray.Length > 4) {
                        islongTable = true;
                    }

                    sb.AddTableHeader(row);

                }

                int j=0;
                sb.Append("\t\t");
                foreach (object cell in row.ItemArray) {
                    string value=GetCellValue(dataSet, cell, islongTable);
                    if (j > 0) {
                        sb.Append(" & ");
                    }
                    sb.Append(value);
                    j++;
                }
                sb.Append(@"\\ \hline").Append("\n");

                i++;
            }
            sb.Append("\t\\end{tabular}\n");
            sb.AppendLine("\\end{center}");
            return sb.ToString();
        }

        private static void AddTableHeader(this StringBuilder sb, DataRow row) {
            int j = 0;
            sb.Append("\t\\begin{tabular}{");
            foreach (object cell in row.ItemArray) {
                if (j == 0) {
                    sb.Append("| ");
                }
                sb.Append(@"p{\q} |");

                j++;
            }
            sb.Append("}\n");
            sb.Append("\t\t\\hline\n");
        }

        private static string GetCellValue(DataSet dataSet, object cell,bool islongTable) {
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
            if (Regex.IsMatch(value, @"^(-?[0-9]{1,}[.]?[0-9]*)$")) {
                if (Config.DecimalPrecision > 0) {
                    var old = value;
                    value = string.Format(Config.DecimalFormat, Double.Parse(value));
                }
                if (islongTable) {
                    value = ToVerticalableNumber(value);
                }
            }

            return value.ToTexEscape();
        }

        private static string ToTexEscape(this string input) {

            var escapes = new char[] {
                '#', '$', '%', '^', '&', '_', '{', '}', '~'
            };

            input = input.Replace("\\", "$\\backslash$");

            input = Regex.Replace(input, @"[#$%^&_{}~]", delegate (Match match) {
                string v = match.ToString();
                return "\\" + v;
            });

            return input;
        }

        private static string ToVerticalableNumber(this string s) {

            if (!Config.SplitNumber) {
                return s;
            }

            var sb = new StringBuilder();
            var i = 0;
            foreach (var c in s) {
                if (i > 0) {
                    sb.Append(' ');
                }
                sb.Append(c);
                i++;
            }
            return sb.ToString();
        }
    }
}