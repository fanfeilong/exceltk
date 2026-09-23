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
            if (Config.PrettyTable) {
                return ToPrettyMd(table, dataSet, insertHeader);
            }
            return ToCompactMd(table, dataSet, insertHeader);
        }

        private static string ToCompactMd(DataTable table, DataSet dataSet, bool insertHeader) {
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

        private static string ToPrettyMd(DataTable table, DataSet dataSet, bool insertHeader) {
            int colCount=table.Columns.Count;
            var rows=new List<string[]>();
            foreach (DataRow row in table.Rows) {
                var cells=new string[colCount];
                for (int c=0; c<colCount; c++) {
                    cells[c]=GetCellValue(dataSet, row.ItemArray[c]);
                }
                rows.Add(cells);
            }
            if (rows.Count==0) {
                return "";
            }

            // BodyHead wraps the first data row in **bold**; account for that in widths.
            var displayRows=new List<string[]>();
            for (int r=0; r<rows.Count; r++) {
                var cells=new string[colCount];
                for (int c=0; c<colCount; c++) {
                    if (r==0 && Config.BodyHead) {
                        cells[c]="**"+rows[r][c]+"**";
                    } else {
                        cells[c]=rows[r][c];
                    }
                }
                displayRows.Add(cells);
            }

            var widths=new int[colCount];
            for (int c=0; c<colCount; c++) {
                // Separator needs at least 3 dashes (":--", "--:", ":--:").
                widths[c]=3;
                foreach (var cells in displayRows) {
                    widths[c]=Math.Max(widths[c], cells[c].Length);
                }
                if (Config.BodyHead && insertHeader) {
                    // Empty header cells still occupy column width.
                    widths[c]=Math.Max(widths[c], 0);
                }
            }

            var sb=new StringBuilder();
            if (Config.BodyHead && insertHeader) {
                AppendPrettyRow(sb, EmptyCells(colCount), widths);
                AppendPrettySeparator(sb, widths);
            }

            for (int r=0; r<displayRows.Count; r++) {
                AppendPrettyRow(sb, displayRows[r], widths);
                if (!Config.BodyHead && r==0 && insertHeader) {
                    AppendPrettySeparator(sb, widths);
                }
            }
            return sb.ToString();
        }

        private static string[] EmptyCells(int colCount) {
            var cells=new string[colCount];
            for (int i=0; i<colCount; i++) {
                cells[i]="";
            }
            return cells;
        }

        private static void AppendPrettyRow(StringBuilder sb, string[] cells, int[] widths) {
            sb.Append("|");
            for (int c=0; c<cells.Length; c++) {
                sb.Append(" ").Append(PadCell(cells[c], widths[c])).Append(" |");
            }
            sb.Append(Environment.NewLine);
        }

        private static void AppendPrettySeparator(StringBuilder sb, int[] widths) {
            sb.Append("|");
            for (int c=0; c<widths.Length; c++) {
                sb.Append(" ").Append(PadSeparator(widths[c])).Append(" |");
            }
            sb.Append(Environment.NewLine);
        }

        private static string PadCell(string value, int width) {
            int pad=Math.Max(0, width-value.Length);
            switch (Config.TableAligin) {
                case "r":
                    return new string(' ', pad)+value;
                case "c": {
                    int left=pad/2;
                    int right=pad-left;
                    return new string(' ', left)+value+new string(' ', right);
                }
                default:
                    return value+new string(' ', pad);
            }
        }

        private static string PadSeparator(int width) {
            // width is content width; separator fills the same visible width.
            int dashCount=Math.Max(3, width);
            switch (Config.TableAligin) {
                case "r":
                    return new string('-', dashCount-1)+":";
                case "c":
                    return ":"+new string('-', Math.Max(1, dashCount-2))+":";
                default:
                    return ":"+new string('-', dashCount-1);
            }
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
