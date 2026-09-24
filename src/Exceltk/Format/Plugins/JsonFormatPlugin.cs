using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.Json;

using Exceltk.Reader;

namespace Exceltk.Format.Plugins {
    public sealed class JsonFormatPlugin : IFormatPlugin {
        public string Name { get { return "json"; } }
        public string FileExtension { get { return "json"; } }
        public string Description { get { return "JSON table (export/import)"; } }
        public bool SupportsExport { get { return true; } }
        public bool SupportsImport { get { return true; } }

        public IEnumerable<FormatArtifact> Export(DataSet dataSet, string sheetFilter) {
            // Precise TableDocument JSON so import round-trips cleanly.
            TableDocument document = TableDocument.FromDataSet(ShrinkCopy(dataSet, sheetFilter));
            if (!string.IsNullOrEmpty(sheetFilter)) {
                foreach (SheetDocument sheet in document.Sheets) {
                    var one = new TableDocument();
                    one.Sheets.Add(sheet);
                    yield return FormatArtifact.Text(sheet.Name, FileExtension, one.ToJson());
                }
            } else if (document.Sheets.Count == 1) {
                yield return FormatArtifact.Text(document.Sheets[0].Name, FileExtension, document.ToJson());
            } else {
                foreach (SheetDocument sheet in document.Sheets) {
                    var one = new TableDocument();
                    one.Sheets.Add(sheet);
                    yield return FormatArtifact.Text(sheet.Name, FileExtension, one.ToJson());
                }
            }
        }

        private static DataSet ShrinkCopy(DataSet dataSet, string sheetFilter) {
            var copy = new DataSet();
            foreach (DataTable table in MarkdownFormatPlugin.FilterSheets(dataSet, sheetFilter)) {
                DataTable snapshot = MarkdownFormatPlugin.Snapshot(table);
                snapshot.Shrink();
                copy.Tables.Add(snapshot);
            }
            return copy;
        }

        public DataSet Import(Stream input, string sourcePath) {
            string text;
            using (var reader = new StreamReader(input, Encoding.UTF8, true, 1024, true)) {
                text = reader.ReadToEnd();
            }
            text = text.Trim();
            if (text.Length == 0) {
                throw new FormatException("Empty JSON import.");
            }

            // Prefer precise ExcelTk document when present.
            if (text.Contains("\"marker\"") && text.Contains(TableDocument.Marker)) {
                try {
                    return TableDocument.FromJson(text).ToDataSet();
                } catch {
                    // fall through to legacy shape
                }
            }

            if (text.StartsWith("[", StringComparison.Ordinal)) {
                return ImportArray(text, sourcePath);
            }
            return ImportObject(text, sourcePath);
        }

        private static DataSet ImportObject(string text, string sourcePath) {
            // Legacy exporter uses single-quoted JS-like JSON; normalize quotes.
            string normalized = NormalizeLegacyJson(text);
            using (JsonDocument doc = JsonDocument.Parse(normalized)) {
                JsonElement root = doc.RootElement;
                string name = root.TryGetProperty("name", out JsonElement nameEl)
                    ? nameEl.GetString()
                    : Path.GetFileNameWithoutExtension(sourcePath ?? "Sheet1");
                var sheet = new SheetDocument { Name = string.IsNullOrEmpty(name) ? "Sheet1" : name };
                if (root.TryGetProperty("rows", out JsonElement rows) && rows.ValueKind == JsonValueKind.Array) {
                    string[] headers = null;
                    foreach (JsonElement row in rows.EnumerateArray()) {
                        if (row.ValueKind != JsonValueKind.Object) {
                            continue;
                        }
                        if (headers == null) {
                            var keys = new List<string>();
                            foreach (JsonProperty prop in row.EnumerateObject()) {
                                keys.Add(prop.Name);
                            }
                            headers = keys.ToArray();
                            sheet.Rows.Add(new List<string>(headers));
                        }
                        var values = new List<string>();
                        foreach (string key in headers) {
                            JsonElement val;
                            values.Add(row.TryGetProperty(key, out val) ? ValueToString(val) : "");
                        }
                        sheet.Rows.Add(values);
                    }
                }
                var dataSet = new DataSet();
                dataSet.Tables.Add(sheet.ToDataTable());
                return dataSet;
            }
        }

        private static DataSet ImportArray(string text, string sourcePath) {
            string normalized = NormalizeLegacyJson(text);
            using (JsonDocument doc = JsonDocument.Parse(normalized)) {
                var dataSet = new DataSet();
                int index = 1;
                foreach (JsonElement el in doc.RootElement.EnumerateArray()) {
                    string name = "Sheet" + index;
                    if (el.ValueKind == JsonValueKind.Object && el.TryGetProperty("name", out JsonElement n)) {
                        name = n.GetString() ?? name;
                    }
                    // Re-serialize single object through object importer.
                    string one = el.GetRawText();
                    DataSet part = ImportObject(one, name);
                    foreach (DataTable table in part.Tables) {
                        if (dataSet.Tables.ContainsTable(table.TableName)) {
                            table.TableName = name + "_" + index;
                        }
                        dataSet.Tables.Add(table);
                    }
                    index++;
                }
                if (dataSet.Tables.Count == 0) {
                    throw new FormatException("No JSON sheets found to import.");
                }
                return dataSet;
            }
        }

        private static string NormalizeLegacyJson(string text) {
            // Convert {'a':'b'} style into {"a":"b"} for System.Text.Json.
            // Only rewrite quotes that are JSON structural, leave escaped content alone via simple pass.
            var sb = new StringBuilder(text.Length);
            bool inString = false;
            char quote = '\0';
            for (int i = 0; i < text.Length; i++) {
                char ch = text[i];
                if (!inString) {
                    if (ch == '\'' || ch == '"') {
                        inString = true;
                        quote = ch;
                        sb.Append('"');
                    } else {
                        sb.Append(ch);
                    }
                    continue;
                }

                if (ch == '\\' && i + 1 < text.Length) {
                    char next = text[i + 1];
                    if (next == quote) {
                        sb.Append('\\').Append('"');
                        i++;
                        continue;
                    }
                    sb.Append(ch).Append(next);
                    i++;
                    continue;
                }
                if (ch == quote) {
                    inString = false;
                    sb.Append('"');
                    continue;
                }
                if (ch == '"' && quote == '\'') {
                    sb.Append('\\').Append('"');
                    continue;
                }
                sb.Append(ch);
            }
            return sb.ToString();
        }

        private static string ValueToString(JsonElement val) {
            switch (val.ValueKind) {
                case JsonValueKind.String:
                    return UnescapeLegacy(val.GetString() ?? "");
                case JsonValueKind.Number:
                case JsonValueKind.True:
                case JsonValueKind.False:
                    return val.ToString();
                case JsonValueKind.Null:
                case JsonValueKind.Undefined:
                    return "";
                default:
                    return val.GetRawText();
            }
        }

        private static string UnescapeLegacy(string value) {
            return value
                .Replace("\\n", "\n")
                .Replace("\\r", "\r")
                .Replace("\\t", "\t");
        }
    }
}
