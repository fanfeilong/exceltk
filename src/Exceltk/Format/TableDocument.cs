using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

using Exceltk.Reader;

namespace Exceltk.Format {
    /// <summary>
    /// Precise, format-neutral workbook payload used for round-trip import/export.
    /// Image plugins embed this inside a marked PNG chunk.
    /// </summary>
    public sealed class TableDocument {
        public const string Marker = "EXCELTK1";
        public const int Version = 1;

        [JsonPropertyName("marker")]
        public string MarkerValue { get; set; } = Marker;

        [JsonPropertyName("version")]
        public int VersionValue { get; set; } = Version;

        [JsonPropertyName("sheets")]
        public List<SheetDocument> Sheets { get; set; } = new List<SheetDocument>();

        public static TableDocument FromDataSet(DataSet dataSet, string sheetFilter = null) {
            var doc = new TableDocument();
            foreach (DataTable table in dataSet.Tables) {
                if (!string.IsNullOrEmpty(sheetFilter)
                    && !string.Equals(table.TableName, sheetFilter, StringComparison.OrdinalIgnoreCase)) {
                    continue;
                }
                doc.Sheets.Add(SheetDocument.FromDataTable(table, dataSet));
            }
            if (!string.IsNullOrEmpty(sheetFilter) && doc.Sheets.Count == 0) {
                throw new ArgumentException("Sheet not found: " + sheetFilter);
            }
            return doc;
        }

        public DataSet ToDataSet() {
            var dataSet = new DataSet();
            foreach (SheetDocument sheet in Sheets) {
                dataSet.Tables.Add(sheet.ToDataTable());
            }
            return dataSet;
        }

        public string ToJson() {
            return JsonSerializer.Serialize(this, JsonOptions());
        }

        public static TableDocument FromJson(string json) {
            if (string.IsNullOrWhiteSpace(json)) {
                throw new FormatException("Empty ExcelTk table document.");
            }
            TableDocument doc = JsonSerializer.Deserialize<TableDocument>(json, JsonOptions());
            if (doc == null) {
                throw new FormatException("Invalid ExcelTk table document.");
            }
            if (!string.Equals(doc.MarkerValue, Marker, StringComparison.Ordinal)) {
                throw new FormatException("ExcelTk marker missing or unsupported: " + doc.MarkerValue);
            }
            if (doc.VersionValue != Version) {
                throw new FormatException("Unsupported ExcelTk document version: " + doc.VersionValue);
            }
            if (doc.Sheets == null) {
                doc.Sheets = new List<SheetDocument>();
            }
            return doc;
        }

        public byte[] ToMarkedBytes() {
            byte[] marker = Encoding.ASCII.GetBytes(Marker);
            byte[] json = Encoding.UTF8.GetBytes(ToJson());
            var payload = new byte[marker.Length + 1 + json.Length];
            Buffer.BlockCopy(marker, 0, payload, 0, marker.Length);
            payload[marker.Length] = 0;
            Buffer.BlockCopy(json, 0, payload, marker.Length + 1, json.Length);
            return payload;
        }

        public static TableDocument FromMarkedBytes(byte[] payload) {
            if (payload == null || payload.Length < Marker.Length + 1) {
                throw new FormatException("Marked payload too short.");
            }
            string magic = Encoding.ASCII.GetString(payload, 0, Marker.Length);
            if (!string.Equals(magic, Marker, StringComparison.Ordinal)) {
                throw new FormatException("ExcelTk image marker not found.");
            }
            if (payload[Marker.Length] != 0) {
                throw new FormatException("ExcelTk image marker is malformed.");
            }
            string json = Encoding.UTF8.GetString(payload, Marker.Length + 1, payload.Length - Marker.Length - 1);
            return FromJson(json);
        }

        private static JsonSerializerOptions JsonOptions() {
            return new JsonSerializerOptions {
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
                WriteIndented = false,
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
            };
        }
    }

    public sealed class SheetDocument {
        [JsonPropertyName("name")]
        public string Name { get; set; }

        [JsonPropertyName("rows")]
        public List<List<string>> Rows { get; set; } = new List<List<string>>();

        [JsonPropertyName("merges")]
        public List<MergeDocument> Merges { get; set; } = new List<MergeDocument>();

        public static SheetDocument FromDataTable(DataTable table, DataSet dataSet) {
            var sheet = new SheetDocument {
                Name = table.TableName ?? ""
            };
            int width = table.Columns.Count;
            foreach (DataRow row in table.Rows) {
                var cells = new List<string>(width);
                object[] items = row.ItemArray ?? Array.Empty<object>();
                for (int c = 0; c < width; c++) {
                    object cell = c < items.Length ? items[c] : null;
                    cells.Add(CellText(dataSet, cell));
                }
                sheet.Rows.Add(cells);
            }
            if (table.Merges != null) {
                foreach (CellMerge merge in table.Merges) {
                    sheet.Merges.Add(new MergeDocument {
                        Row = merge.Row,
                        Col = merge.Col,
                        RowSpan = merge.RowSpan,
                        ColSpan = merge.ColSpan
                    });
                }
            }
            return sheet;
        }

        public DataTable ToDataTable() {
            string name = string.IsNullOrEmpty(Name) ? "Sheet1" : Name;
            var table = new DataTable(name);
            int width = 0;
            foreach (List<string> row in Rows) {
                if (row != null) {
                    width = Math.Max(width, row.Count);
                }
            }
            for (int i = 0; i < width; i++) {
                table.Columns.Add(i.ToString(CultureInfo.InvariantCulture), typeof(object));
            }
            foreach (List<string> row in Rows) {
                var cells = new object[width];
                for (int c = 0; c < width; c++) {
                    string value = row != null && c < row.Count ? row[c] : "";
                    cells[c] = new XlsCell(value ?? "");
                }
                table.Rows.Add(cells);
            }
            if (Merges != null) {
                foreach (MergeDocument merge in Merges) {
                    table.Merges.Add(new CellMerge {
                        Row = merge.Row,
                        Col = merge.Col,
                        RowSpan = merge.RowSpan,
                        ColSpan = merge.ColSpan
                    });
                }
            }
            return table;
        }

        private static string CellText(DataSet dataSet, object cell) {
            if (cell == null) {
                return "";
            }
            var xlsCell = cell as XlsCell;
            if (xlsCell != null) {
                return xlsCell.GetMarkDownText(dataSet) ?? "";
            }
            return cell.ToString() ?? "";
        }
    }

    public sealed class MergeDocument {
        [JsonPropertyName("row")]
        public int Row { get; set; }

        [JsonPropertyName("col")]
        public int Col { get; set; }

        [JsonPropertyName("rowSpan")]
        public int RowSpan { get; set; }

        [JsonPropertyName("colSpan")]
        public int ColSpan { get; set; }
    }
}
