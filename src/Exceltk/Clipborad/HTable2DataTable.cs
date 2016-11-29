#if OS_WINDOWS
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Exceltk;

namespace Exceltk.Clipborad {
    /// <summary>
    /// http://blog.hypercomplex.co.uk/index.php/2010/05/parsing-html-tables-into-system-data-datatable/
    /// HtmlTableParser parses the contents of an html string into a System.Data DataSet or DataTable.
    /// </summary>
    public static class HTable2DataTable {
        private const RegexOptions ExpressionOptions = RegexOptions.Singleline | RegexOptions.Multiline | RegexOptions.IgnoreCase;

        private const string CommentPattern = "<!--(.*?)-->";
        private const string TablePattern = "<table[^>]*>(.*?)</table>";
        private const string HeaderPattern = "<th[^>]*>(.*?)</th>";
        private const string RowPattern = "<tr[^>]*>(.*?)</tr>";
        private const string CellPattern = "<td[^>]*>(.*?)</td>";
        private const string HyperLinkPattern = "<a\\s*href\\s*=\\s*\"(.*?)\"\\s*>(.*?)</a>";

        private const string CharsetPattern = "<meta [^>]*charset=(.*?)\">";

        /// <summary>
        /// Given an HTML string containing n table tables, parse them into a DataSet containing n DataTables.
        /// </summary>
        /// <param name="htmlObj">An HTML string containing n HTML tables</param>
        /// <returns>A DataSet containing a DataTable for each HTML table in the input HTML</returns>
        public static DataSet ParseDataSet(this object htmlObj) {
            
            var ms = htmlObj as MemoryStream;
            Debug.Assert(ms!=null);

            ms.Position = 0;
            var bytes = new byte[ms.Length];
            ms.Read(bytes, 0, (int)ms.Length);

            string html = Encoding.UTF8.GetString(bytes);

            var charsetMacth = Regex.Match(html, CharsetPattern);
            
            if (charsetMacth.Captures.Count > 0) {
                var charset = charsetMacth.Groups[1].Value;
                if (charset != "utf-8") {
                    Console.WriteLine("ERROR: Invalid Encoding...");
                    return null;
                }
            }

            var dataSet = new DataSet();
            var tableMatches = Regex.Matches(
                WithoutComments(html),
                TablePattern,
                ExpressionOptions);

            foreach (Match tableMatch in tableMatches) {
                dataSet.Tables.Add(ParseTable(tableMatch.Value));
            }

            return dataSet;
        }

        /// <summary>
        /// Given an HTML string containing a single table, parse that table to form a DataTable.
        /// </summary>
        /// <param name="tableHtml">An HTML string containing a single HTML table</param>
        /// <returns>A DataTable which matches the input HTML table</returns>
        public static DataTable ParseTable(string tableHtml) {
            string tableHtmlWithoutComments = WithoutComments(tableHtml);

            var dataTable = new DataTable("");

            var rowMatches = Regex.Matches(
                tableHtmlWithoutComments,
                RowPattern,
                ExpressionOptions);

            dataTable.Columns.AddRange(tableHtmlWithoutComments.Contains("<th")
                                           ? ParseColumns(tableHtml)
                                           : GenerateColumns(rowMatches));

            ParseRows(rowMatches, dataTable);

            return dataTable;
        }

        /// <summary>
        /// Strip comments from an HTML stirng
        /// </summary>
        /// <param name="html">An HTML string potentially containing comments</param>
        /// <returns>The input HTML string with comments removed</returns>
        private static string WithoutComments(string html) {
            return Regex.Replace(html, CommentPattern, string.Empty, ExpressionOptions);
        }

        /// <summary>
        /// Add a row to the input DataTable for each row match in the input MatchCollection
        /// </summary>
        /// <param name="rowMatches">A collection of all the rows to add to the DataTable</param>
        /// <param name="dataTable">The DataTable to which we add rows</param>
        private static void ParseRows(IEnumerable rowMatches, DataTable dataTable) {
            foreach (Match rowMatch in rowMatches) {
                // if the row contains header tags don't use it - it is a header not a row
                if (!rowMatch.Value.Contains("<th")) {
                    DataRow dataRow = dataTable.NewRow();
                    var rowArray = new List<XlsCell>();
                    MatchCollection cellMatches = Regex.Matches(
                        rowMatch.Value,
                        CellPattern,
                        ExpressionOptions);

                    for (int columnIndex = 0; columnIndex < cellMatches.Count; columnIndex++) {
                        var cellValue = cellMatches[columnIndex].Groups[1].ToString();
                        cellValue = cellValue.RemoveSpan();

                        var hyperLinkMach = Regex.Match(cellValue, HyperLinkPattern);
                        if (hyperLinkMach.Captures.Count >0) {
                            var xlsCell = new XlsCell(hyperLinkMach.Groups[2].Value);
                            xlsCell.SetHyperLink(hyperLinkMach.Groups[1].Value);
                            rowArray.Add(xlsCell);
                            Console.WriteLine("have hyperlink");
                        } else {
                            rowArray.Add(new XlsCell(cellValue));
                        }
                    }
                    XlsCell[] cells = rowArray.ToArray();
                    dataRow.ItemArray = cells;
                    dataTable.Rows.Add(dataRow);
                }
            }
        }

        /// <summary>
        /// Given a string containing an HTML table, parse the header cells to create a set of DataColumns
        /// which define the columns in a DataTable.
        /// </summary>
        /// <param name="tableHtml">An HTML string containing a single HTML table</param>
        /// <returns>A set of DataColumns based on the HTML table header cells</returns>
        private static DataColumn[] ParseColumns(string tableHtml) {
            MatchCollection headerMatches = Regex.Matches(
                tableHtml,
                HeaderPattern,
                ExpressionOptions);

            return (from Match headerMatch in headerMatches
                    select new DataColumn(headerMatch.Groups[1].ToString(),typeof(Object))).ToArray();
        }

        /// <summary>
        /// For tables which do not specify header cells we must generate DataColumns based on the number
        /// of cells in a row (we assume all rows have the same number of cells).
        /// </summary>
        /// <param name="rowMatches">A collection of all the rows in the HTML table we wish to generate columns for</param>
        /// <returns>A set of DataColumns based on the number of celss in the first row of the input HTML table</returns>
        private static DataColumn[] GenerateColumns(MatchCollection rowMatches) {
            int columnCount = Regex.Matches(
                rowMatches[0].ToString(),
                CellPattern,
                ExpressionOptions).Count;

            return (from index in Enumerable.Range(0, columnCount)
                    select new DataColumn("Column " + Convert.ToString(index),typeof(Object))).ToArray();
        }

        private static string RemoveSpan(this string text) {

            const string bspan = "<span";
            const string espan = "</span>";

            int start = text.IndexOf(bspan, System.StringComparison.Ordinal);
            int end = text.IndexOf(espan, System.StringComparison.Ordinal);
            while (start >= 0 && end >= 0) {
                text = text.Remove(start, end - start+espan.Length);
                start = text.IndexOf(bspan, System.StringComparison.Ordinal);
                end = text.IndexOf(espan, System.StringComparison.Ordinal);
            }
            return text;
        }
    }
}
#endif
