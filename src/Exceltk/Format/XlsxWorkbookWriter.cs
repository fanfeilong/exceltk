using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

using Exceltk.Reader;
using ICSharpCode.SharpZipLib.Zip;

namespace Exceltk.Format {
    /// <summary>
    /// Minimal OpenXML (.xlsx) writer used by format-plugin import.
    /// </summary>
    public static class XlsxWorkbookWriter {
        public static void Write(DataSet dataSet, string path) {
            if (dataSet == null) {
                throw new ArgumentNullException(nameof(dataSet));
            }
            if (string.IsNullOrEmpty(path)) {
                throw new ArgumentNullException(nameof(path));
            }

            string dir = Path.GetDirectoryName(path);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir)) {
                Directory.CreateDirectory(dir);
            }

            using (FileStream fs = File.Create(path))
            using (var zip = new ZipOutputStream(fs)) {
                zip.SetLevel(6);
                zip.IsStreamOwner = false;

                var sheets = new List<DataTable>();
                foreach (DataTable table in dataSet.Tables) {
                    sheets.Add(table);
                }
                if (sheets.Count == 0) {
                    sheets.Add(new DataTable("Sheet1"));
                }

                var shared = new SharedStringTable();
                var sheetXml = new List<string>();
                for (int i = 0; i < sheets.Count; i++) {
                    sheetXml.Add(BuildSheetXml(sheets[i], shared));
                }

                WriteEntry(zip, "[Content_Types].xml", BuildContentTypes(sheets.Count));
                WriteEntry(zip, "_rels/.rels", BuildRootRels());
                WriteEntry(zip, "xl/workbook.xml", BuildWorkbook(sheets));
                WriteEntry(zip, "xl/_rels/workbook.xml.rels", BuildWorkbookRels(sheets.Count));
                WriteEntry(zip, "xl/styles.xml", BuildStyles());
                WriteEntry(zip, "xl/sharedStrings.xml", shared.ToXml());
                for (int i = 0; i < sheets.Count; i++) {
                    WriteEntry(zip, "xl/worksheets/sheet" + (i + 1) + ".xml", sheetXml[i]);
                }

                zip.Finish();
            }
        }

        private static void WriteEntry(ZipOutputStream zip, string name, string content) {
            var entry = new ZipEntry(name);
            entry.DateTime = DateTime.Now;
            zip.PutNextEntry(entry);
            byte[] bytes = Encoding.UTF8.GetBytes(content);
            zip.Write(bytes, 0, bytes.Length);
            zip.CloseEntry();
        }

        private static string BuildContentTypes(int sheetCount) {
            var sb = new StringBuilder();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">");
            sb.Append("<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>");
            sb.Append("<Default Extension=\"xml\" ContentType=\"application/xml\"/>");
            sb.Append("<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>");
            sb.Append("<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>");
            sb.Append("<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>");
            for (int i = 1; i <= sheetCount; i++) {
                sb.Append("<Override PartName=\"/xl/worksheets/sheet")
                    .Append(i)
                    .Append(".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>");
            }
            sb.Append("</Types>");
            return sb.ToString();
        }

        private static string BuildRootRels() {
            return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>"
                + "</Relationships>";
        }

        private static string BuildWorkbook(List<DataTable> sheets) {
            var sb = new StringBuilder();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append("<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" ");
            sb.Append("xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
            sb.Append("<sheets>");
            for (int i = 0; i < sheets.Count; i++) {
                string name = string.IsNullOrEmpty(sheets[i].TableName) ? ("Sheet" + (i + 1)) : sheets[i].TableName;
                sb.Append("<sheet name=\"")
                    .Append(XmlEscape(SanitizeSheetName(name)))
                    .Append("\" sheetId=\"")
                    .Append(i + 1)
                    .Append("\" r:id=\"rId")
                    .Append(i + 1)
                    .Append("\"/>");
            }
            sb.Append("</sheets></workbook>");
            return sb.ToString();
        }

        private static string BuildWorkbookRels(int sheetCount) {
            var sb = new StringBuilder();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
            for (int i = 1; i <= sheetCount; i++) {
                sb.Append("<Relationship Id=\"rId")
                    .Append(i)
                    .Append("\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet")
                    .Append(i)
                    .Append(".xml\"/>");
            }
            int next = sheetCount + 1;
            sb.Append("<Relationship Id=\"rId")
                .Append(next)
                .Append("\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>");
            next++;
            sb.Append("<Relationship Id=\"rId")
                .Append(next)
                .Append("\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>");
            sb.Append("</Relationships>");
            return sb.ToString();
        }

        private static string BuildStyles() {
            return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                + "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
                + "<fonts count=\"1\"><font><sz val=\"11\"/><color theme=\"1\"/><name val=\"Calibri\"/><family val=\"2\"/></font></fonts>"
                + "<fills count=\"1\"><fill><patternFill patternType=\"none\"/></fill></fills>"
                + "<borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders>"
                + "<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>"
                + "<cellXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/></cellXfs>"
                + "</styleSheet>";
        }

        private static string BuildSheetXml(DataTable table, SharedStringTable shared) {
            int rowCount = table.Rows.Count;
            int colCount = Math.Max(1, table.Columns.Count);
            var sb = new StringBuilder();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
            sb.Append("<dimension ref=\"A1:")
                .Append(ColumnName(colCount - 1))
                .Append(Math.Max(1, rowCount))
                .Append("\"/>");
            sb.Append("<sheetData>");
            for (int r = 0; r < rowCount; r++) {
                object[] items = table.Rows[r].ItemArray ?? Array.Empty<object>();
                sb.Append("<row r=\"").Append(r + 1).Append("\">");
                for (int c = 0; c < colCount; c++) {
                    string text = CellValue(c < items.Length ? items[c] : null);
                    int index = shared.GetIndex(text);
                    string refName = ColumnName(c) + (r + 1).ToString(CultureInfo.InvariantCulture);
                    sb.Append("<c r=\"")
                        .Append(refName)
                        .Append("\" t=\"s\"><v>")
                        .Append(index)
                        .Append("</v></c>");
                }
                sb.Append("</row>");
            }
            sb.Append("</sheetData>");

            if (table.Merges != null && table.Merges.Count > 0) {
                sb.Append("<mergeCells count=\"").Append(table.Merges.Count).Append("\">");
                foreach (CellMerge merge in table.Merges) {
                    string a = ColumnName(merge.Col) + (merge.Row + 1);
                    string b = ColumnName(merge.Col + Math.Max(1, merge.ColSpan) - 1)
                        + (merge.Row + Math.Max(1, merge.RowSpan));
                    sb.Append("<mergeCell ref=\"").Append(a).Append(":").Append(b).Append("\"/>");
                }
                sb.Append("</mergeCells>");
            }

            sb.Append("</worksheet>");
            return sb.ToString();
        }

        private static string CellValue(object cell) {
            if (cell == null) {
                return "";
            }
            var xls = cell as XlsCell;
            if (xls != null) {
                return xls.Value == null ? "" : xls.Value.ToString();
            }
            return cell.ToString() ?? "";
        }

        private static string ColumnName(int zeroBased) {
            int n = zeroBased + 1;
            var sb = new StringBuilder();
            while (n > 0) {
                int rem = (n - 1) % 26;
                sb.Insert(0, (char)('A' + rem));
                n = (n - 1) / 26;
            }
            return sb.ToString();
        }

        private static string SanitizeSheetName(string name) {
            if (string.IsNullOrEmpty(name)) {
                return "Sheet1";
            }
            var cleaned = name.Replace("\\", " ").Replace("/", " ").Replace("?", " ")
                .Replace("*", " ").Replace("[", " ").Replace("]", " ").Replace(":", " ");
            if (cleaned.Length > 31) {
                cleaned = cleaned.Substring(0, 31);
            }
            return string.IsNullOrWhiteSpace(cleaned) ? "Sheet1" : cleaned;
        }

        private static string XmlEscape(string value) {
            if (string.IsNullOrEmpty(value)) {
                return "";
            }
            return value
                .Replace("&", "&amp;")
                .Replace("<", "&lt;")
                .Replace(">", "&gt;")
                .Replace("\"", "&quot;")
                .Replace("'", "&apos;");
        }

        private sealed class SharedStringTable {
            private readonly Dictionary<string, int> _index = new Dictionary<string, int>(StringComparer.Ordinal);
            private readonly List<string> _values = new List<string>();

            public int GetIndex(string value) {
                value = value ?? "";
                int existing;
                if (_index.TryGetValue(value, out existing)) {
                    return existing;
                }
                int idx = _values.Count;
                _values.Add(value);
                _index[value] = idx;
                return idx;
            }

            public string ToXml() {
                var sb = new StringBuilder();
                sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
                sb.Append("<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"")
                    .Append(_values.Count)
                    .Append("\" uniqueCount=\"")
                    .Append(_values.Count)
                    .Append("\">");
                foreach (string value in _values) {
                    sb.Append("<si><t xml:space=\"preserve\">")
                        .Append(XmlEscape(value))
                        .Append("</t></si>");
                }
                sb.Append("</sst>");
                return sb.ToString();
            }
        }
    }
}
