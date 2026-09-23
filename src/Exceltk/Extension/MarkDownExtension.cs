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
            if (Config.MultiMarkdown) {
                ApplyMerges(table);
                return ToMultiMarkdownHtml(table, dataSet, insertHeader);
            }
            if (Config.PrettyTable) {
                return ToPrettyMd(table, dataSet, insertHeader);
            }
            return ToCompactMd(table, dataSet, insertHeader);
        }

        /// <summary>
        /// Apply Excel merge regions onto cells (rowspan/colspan + covered markers).
        /// Coordinates are clipped to the post-Shrink table size.
        /// </summary>
        private static void ApplyMerges(DataTable table) {
            if (table.Merges==null || table.Merges.Count==0) {
                return;
            }
            int rowCount=table.Rows.Count;
            int colCount=table.Columns.Count;
            foreach (CellMerge merge in table.Merges) {
                if (merge.Row<0 || merge.Col<0 || merge.Row>=rowCount || merge.Col>=colCount) {
                    continue;
                }
                int rowSpan=Math.Min(merge.RowSpan, rowCount-merge.Row);
                int colSpan=Math.Min(merge.ColSpan, colCount-merge.Col);
                if (rowSpan<=1 && colSpan<=1) {
                    continue;
                }

                object anchorObj=table.Rows[merge.Row][merge.Col];
                XlsCell anchor=anchorObj as XlsCell;
                if (anchor==null) {
                    anchor=new XlsCell(anchorObj??"");
                    table.Rows[merge.Row][merge.Col]=anchor;
                }
                anchor.RowSpan=rowSpan;
                anchor.ColSpan=colSpan;
                anchor.IsMergeCovered=false;

                for (int r=merge.Row; r<merge.Row+rowSpan; r++) {
                    for (int c=merge.Col; c<merge.Col+colSpan; c++) {
                        if (r==merge.Row && c==merge.Col) {
                            continue;
                        }
                        object coveredObj=table.Rows[r][c];
                        XlsCell covered=coveredObj as XlsCell;
                        if (covered==null) {
                            covered=new XlsCell(coveredObj??"");
                            table.Rows[r][c]=covered;
                        }
                        covered.IsMergeCovered=true;
                        covered.RowSpan=1;
                        covered.ColSpan=1;
                    }
                }
            }
        }

        /// <summary>
        /// MultiMarkdown-friendly HTML table with rowspan/colspan for Excel merged cells.
        /// Pipe Markdown cannot express rowspan; HTML tables are valid in MultiMarkdown documents.
        /// </summary>
        private static string ToMultiMarkdownHtml(DataTable table, DataSet dataSet, bool insertHeader) {
            var sb=new StringBuilder();
            string align=HtmlAlignAttribute();
            sb.Append("<table>").Append(Environment.NewLine);

            int startRow=0;
            if (!Config.BodyHead && insertHeader && table.Rows.Count>0) {
                sb.Append("<thead>").Append(Environment.NewLine);
                AppendHtmlRow(sb, table, dataSet, 0, "th", false, align);
                sb.Append("</thead>").Append(Environment.NewLine);
                startRow=1;
            }

            sb.Append("<tbody>").Append(Environment.NewLine);
            if (Config.BodyHead && insertHeader && table.Rows.Count>0) {
                AppendHtmlRow(sb, table, dataSet, 0, "td", true, align);
                startRow=1;
            }
            for (int r=startRow; r<table.Rows.Count; r++) {
                AppendHtmlRow(sb, table, dataSet, r, "td", false, align);
            }
            sb.Append("</tbody>").Append(Environment.NewLine);
            sb.Append("</table>").Append(Environment.NewLine);
            return sb.ToString();
        }

        private static void AppendHtmlRow(StringBuilder sb, DataTable table, DataSet dataSet, int rowIndex, string tag, bool strong, string align) {
            DataRow row=table.Rows[rowIndex];
            sb.Append("<tr>");
            for (int c=0; c<row.ItemArray.Length; c++) {
                object cellObj=row.ItemArray[c];
                var xlsCell=cellObj as XlsCell;
                if (xlsCell!=null && xlsCell.IsMergeCovered) {
                    continue;
                }
                int rowSpan=xlsCell!=null ? Math.Max(1, xlsCell.RowSpan) : 1;
                int colSpan=xlsCell!=null ? Math.Max(1, xlsCell.ColSpan) : 1;
                sb.Append("<").Append(tag).Append(align);
                if (rowSpan>1) {
                    sb.Append(" rowspan=\"").Append(rowSpan).Append("\"");
                }
                if (colSpan>1) {
                    sb.Append(" colspan=\"").Append(colSpan).Append("\"");
                }
                sb.Append(">");
                string value=GetHtmlCellValue(dataSet, cellObj);
                if (strong) {
                    sb.Append("<strong>").Append(value).Append("</strong>");
                } else {
                    sb.Append(value);
                }
                sb.Append("</").Append(tag).Append(">");
            }
            sb.Append("</tr>").Append(Environment.NewLine);
        }

        private static string HtmlAlignAttribute() {
            switch (Config.TableAligin) {
                case "r": return " align=\"right\"";
                case "c": return " align=\"center\"";
                default: return " align=\"left\"";
            }
        }

        private static string GetHtmlCellValue(DataSet dataSet, object cell) {
            if (cell==null) {
                return "";
            }
            var xlsCell=cell as XlsCell;
            if (xlsCell!=null && xlsCell.IsHyperLink) {
                string text=EscapeHtml(ApplyDecimalPrecision(xlsCell.Value!=null ? xlsCell.Value.ToString() : ""));
                text=Regex.Replace(text, @"\r\n?|\n", "<br/>");
                string url=EscapeHtmlAttribute(xlsCell.HyperLink??"");
                return string.Format("<a href=\"{0}\">{1}</a>", url, text);
            }

            string value;
            if (xlsCell!=null) {
                value=xlsCell.GetMarkDownText(dataSet);
            } else {
                value=cell.ToString();
            }

            value=ApplyDecimalPrecision(value);

            // Convert markdown links from hyperlink resolution into HTML anchors.
            Match mdLink=Regex.Match(value, @"^\[([^\]]*)\]\(([^)]*)\)$");
            if (mdLink.Success) {
                return string.Format("<a href=\"{0}\">{1}</a>",
                    EscapeHtmlAttribute(mdLink.Groups[2].Value),
                    EscapeHtml(mdLink.Groups[1].Value));
            }

            value=EscapeHtml(value);
            value=Regex.Replace(value, @"\r\n?|\n", "<br/>");
            return value;
        }

        private static string ApplyDecimalPrecision(string value) {
            if (!Config.HasDecimalPrecision) {
                return value;
            }
            if (Regex.IsMatch(value, @"^(-?[0-9]{1,}[.][0-9]*)$")) {
                if (Config.DecimalPrecision>0) {
                    return string.Format(Config.DecimalFormat, Double.Parse(value));
                }
                return ((int)Double.Parse(value)).ToString();
            }
            return value;
        }

        private static string EscapeHtml(string value) {
            if (string.IsNullOrEmpty(value)) {
                return "";
            }
            return value
                .Replace("&", "&amp;")
                .Replace("<", "&lt;")
                .Replace(">", "&gt;")
                .Replace("\"", "&quot;");
        }

        private static string EscapeHtmlAttribute(string value) {
            return EscapeHtml(value).Replace("'", "&#39;");
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
