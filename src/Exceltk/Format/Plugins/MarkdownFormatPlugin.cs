using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

using Exceltk.Reader;

namespace Exceltk.Format.Plugins {
    public sealed class MarkdownFormatPlugin : IFormatPlugin {
        public string Name { get { return "md"; } }
        public string FileExtension { get { return "md"; } }
        public string Description { get { return "Markdown table (export/import)"; } }
        public bool SupportsExport { get { return true; } }
        public bool SupportsImport { get { return true; } }

        public IEnumerable<FormatArtifact> Export(DataSet dataSet, string sheetFilter) {
            foreach (DataTable table in FilterSheets(dataSet, sheetFilter)) {
                DataTable snapshot = Snapshot(table);
                snapshot.Shrink();
                var one = new DataSet();
                one.Tables.Add(snapshot);
                string marker = EmbedMarkerComment(TableDocument.FromDataSet(one));
                string body = table.ToMd(dataSet);
                yield return FormatArtifact.Text(table.TableName, FileExtension, marker + body);
            }
        }

        internal static DataTable Snapshot(DataTable table) {
            DataTable cloned = table.Clone();
            cloned.Merges = table.Merges != null
                ? new List<CellMerge>(table.Merges)
                : new List<CellMerge>();
            return cloned;
        }

        public DataSet Import(Stream input, string sourcePath) {
            string text;
            using (var reader = new StreamReader(input, Encoding.UTF8, true, 1024, true)) {
                text = reader.ReadToEnd();
            }

            TableDocument marked;
            if (TryExtractMarkedDocument(text, out marked)) {
                return marked.ToDataSet();
            }

            var sheets = ParseMarkdownTables(text);
            if (sheets.Count == 0) {
                throw new FormatException("No markdown table found to import.");
            }

            string baseName = string.IsNullOrEmpty(sourcePath)
                ? "Sheet1"
                : Path.GetFileNameWithoutExtension(sourcePath);
            var dataSet = new DataSet();
            for (int i = 0; i < sheets.Count; i++) {
                string name = sheets.Count == 1 ? baseName : (baseName + "_" + (i + 1));
                sheets[i].Name = name;
                dataSet.Tables.Add(sheets[i].ToDataTable());
            }
            return dataSet;
        }

        /// <summary>
        /// Optional precise payload comment: &lt;!-- EXCELTK1 {...json...} --&gt;
        /// </summary>
        public static string EmbedMarkerComment(TableDocument document) {
            return "<!-- " + TableDocument.Marker + " " + document.ToJson() + " -->" + Environment.NewLine;
        }

        private static bool TryExtractMarkedDocument(string text, out TableDocument document) {
            document = null;
            const string open = "<!--";
            const string close = "-->";
            int from = 0;
            while (true) {
                int start = text.IndexOf(open, from, StringComparison.Ordinal);
                if (start < 0) {
                    return false;
                }
                int end = text.IndexOf(close, start + open.Length, StringComparison.Ordinal);
                if (end < 0) {
                    return false;
                }
                string body = text.Substring(start + open.Length, end - start - open.Length).Trim();
                if (body.StartsWith(TableDocument.Marker, StringComparison.Ordinal)) {
                    string json = body.Substring(TableDocument.Marker.Length).Trim();
                    try {
                        document = TableDocument.FromJson(json);
                        return true;
                    } catch {
                        // keep scanning
                    }
                }
                from = end + close.Length;
            }
        }

        private static List<SheetDocument> ParseMarkdownTables(string text) {
            var sheets = new List<SheetDocument>();
            var lines = text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
            var current = new List<List<string>>();
            foreach (string raw in lines) {
                string line = raw.Trim();
                if (line.StartsWith("|", StringComparison.Ordinal) && line.EndsWith("|", StringComparison.Ordinal)) {
                    if (IsSeparatorRow(line)) {
                        continue;
                    }
                    current.Add(SplitRow(line));
                } else if (current.Count > 0) {
                    sheets.Add(new SheetDocument { Rows = current });
                    current = new List<List<string>>();
                }
            }
            if (current.Count > 0) {
                sheets.Add(new SheetDocument { Rows = current });
            }
            return sheets;
        }

        private static bool IsSeparatorRow(string line) {
            string inner = line.Trim('|');
            return Regex.IsMatch(inner, @"^[\s|:\-]+$");
        }

        private static List<string> SplitRow(string line) {
            string trimmed = line.Trim();
            if (trimmed.StartsWith("|", StringComparison.Ordinal)) {
                trimmed = trimmed.Substring(1);
            }
            if (trimmed.EndsWith("|", StringComparison.Ordinal)) {
                trimmed = trimmed.Substring(0, trimmed.Length - 1);
            }
            var cells = new List<string>();
            var sb = new StringBuilder();
            bool escape = false;
            foreach (char ch in trimmed) {
                if (escape) {
                    sb.Append(ch);
                    escape = false;
                    continue;
                }
                if (ch == '\\') {
                    escape = true;
                    continue;
                }
                if (ch == '|') {
                    cells.Add(CleanCell(sb.ToString()));
                    sb.Clear();
                    continue;
                }
                sb.Append(ch);
            }
            cells.Add(CleanCell(sb.ToString()));
            return cells;
        }

        private static string CleanCell(string value) {
            value = (value ?? "").Trim();
            value = value.Replace("<br/>", "\n").Replace("<br />", "\n");
            if (value.StartsWith("**", StringComparison.Ordinal) && value.EndsWith("**", StringComparison.Ordinal) && value.Length >= 4) {
                value = value.Substring(2, value.Length - 4);
            }
            return value;
        }

        internal static IEnumerable<DataTable> FilterSheets(DataSet dataSet, string sheetFilter) {
            if (string.IsNullOrEmpty(sheetFilter)) {
                foreach (DataTable table in dataSet.Tables) {
                    yield return table;
                }
                yield break;
            }
            if (dataSet.Tables.ContainsTable(sheetFilter)) {
                yield return dataSet.Tables[sheetFilter];
                yield break;
            }
            if (dataSet.Tables.Count == 1) {
                yield return dataSet.Tables[0];
                yield break;
            }
            throw new ArgumentException("Sheet not found: " + sheetFilter);
        }
    }
}
