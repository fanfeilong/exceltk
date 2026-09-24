using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

using Exceltk.Reader;

namespace Exceltk.Format.Plugins {
    public sealed class TexFormatPlugin : IFormatPlugin {
        public string Name { get { return "tex"; } }
        public string FileExtension { get { return "tex"; } }
        public string Description { get { return "TeX tabular (export/import)"; } }
        public bool SupportsExport { get { return true; } }
        public bool SupportsImport { get { return true; } }

        public IEnumerable<FormatArtifact> Export(DataSet dataSet, string sheetFilter) {
            foreach (DataTable table in MarkdownFormatPlugin.FilterSheets(dataSet, sheetFilter)) {
                DataTable snapshot = MarkdownFormatPlugin.Snapshot(table);
                snapshot.Shrink();
                var one = new DataSet();
                one.Tables.Add(snapshot);
                string marker = "% " + TableDocument.Marker + " " + TableDocument.FromDataSet(one).ToJson()
                    + Environment.NewLine;
                string body = table.ToTex(dataSet);
                yield return FormatArtifact.Text(table.TableName, FileExtension, marker + body);
            }
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

            var rows = ParseTabularRows(text);
            if (rows.Count == 0) {
                throw new FormatException("No TeX tabular rows found to import.");
            }

            string name = string.IsNullOrEmpty(sourcePath)
                ? "Sheet1"
                : Path.GetFileNameWithoutExtension(sourcePath);
            var sheet = new SheetDocument { Name = name, Rows = rows };
            var dataSet = new DataSet();
            dataSet.Tables.Add(sheet.ToDataTable());
            return dataSet;
        }

        private static bool TryExtractMarkedDocument(string text, out TableDocument document) {
            document = null;
            using (var reader = new StringReader(text)) {
                string line;
                while ((line = reader.ReadLine()) != null) {
                    string trimmed = line.TrimStart();
                    if (!trimmed.StartsWith("%", StringComparison.Ordinal)) {
                        continue;
                    }
                    trimmed = trimmed.Substring(1).TrimStart();
                    if (!trimmed.StartsWith(TableDocument.Marker, StringComparison.Ordinal)) {
                        continue;
                    }
                    string json = trimmed.Substring(TableDocument.Marker.Length).Trim();
                    try {
                        document = TableDocument.FromJson(json);
                        return true;
                    } catch {
                        // continue
                    }
                }
            }

            // Fallback: HTML-style comment if mixed into TeX output.
            const string open = "<!--";
            const string close = "-->";
            int start = text.IndexOf(open, StringComparison.Ordinal);
            if (start >= 0) {
                int end = text.IndexOf(close, start + open.Length, StringComparison.Ordinal);
                if (end > start) {
                    string body = text.Substring(start + open.Length, end - start - open.Length).Trim();
                    if (body.StartsWith(TableDocument.Marker, StringComparison.Ordinal)) {
                        try {
                            document = TableDocument.FromJson(body.Substring(TableDocument.Marker.Length).Trim());
                            return true;
                        } catch {
                            return false;
                        }
                    }
                }
            }
            return false;
        }

        private static List<List<string>> ParseTabularRows(string text) {
            var rows = new List<List<string>>();
            foreach (Match m in Regex.Matches(text, @"^(?<row>.+?)\\\\\s*(?:\\hline)?\s*$",
                RegexOptions.Multiline)) {
                string rowText = m.Groups["row"].Value.Trim();
                if (rowText.StartsWith("\\", StringComparison.Ordinal)
                    && !rowText.Contains("&")) {
                    continue;
                }
                if (rowText.Contains("begin{") || rowText.Contains("end{") || rowText.Contains("setlength")) {
                    continue;
                }
                string[] parts = rowText.Split('&');
                var cells = new List<string>();
                foreach (string part in parts) {
                    cells.Add(CleanTexCell(part));
                }
                if (cells.Count > 0) {
                    rows.Add(cells);
                }
            }
            return rows;
        }

        private static string CleanTexCell(string value) {
            value = (value ?? "").Trim();
            value = value.Replace(@"\#", "#")
                .Replace(@"\$", "$")
                .Replace(@"\%", "%")
                .Replace(@"\&", "&")
                .Replace(@"\_", "_")
                .Replace(@"\{", "{")
                .Replace(@"\}", "}")
                .Replace(@"~", " ");
            value = Regex.Replace(value, @"\\[a-zA-Z]+\s*", "");
            return value.Trim();
        }
    }
}
