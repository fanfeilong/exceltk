using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Exceltk.Reader {
    /// <summary>
    /// Minimal RFC 4180-style CSV reader that builds the project's DataSet model.
    /// </summary>
    public static class CsvReader {
        public static DataSet AsDataSet(string path) {
            var dataSet=new DataSet();
            var table=new DataTable("");
            using (var reader=new StreamReader(path, Encoding.UTF8, true)) {
                string line;
                int columnCount=0;
                while ((line=reader.ReadLine())!=null) {
                    // Keep trailing empty rows out; leading/middle empty rows become blank cells.
                    if (line.Length==0 && columnCount==0) {
                        continue;
                    }
                    var fields=ParseLine(line);
                    if (columnCount==0) {
                        columnCount=Math.Max(fields.Count, 1);
                        for (int i=0; i<columnCount; i++) {
                            table.Columns.Add("Column"+i);
                        }
                    }
                    if (fields.Count>columnCount) {
                        for (int i=columnCount; i<fields.Count; i++) {
                            table.Columns.Add("Column"+i);
                        }
                        // Pad earlier rows so ItemArray lengths stay consistent.
                        for (int r=0; r<table.Rows.Count; r++) {
                            var old=table.Rows[r].ItemArray;
                            var padded=new object[fields.Count];
                            Array.Copy(old, padded, old.Length);
                            for (int i=old.Length; i<fields.Count; i++) {
                                padded[i]="";
                            }
                            table.Rows[r].ItemArray=padded;
                        }
                        columnCount=fields.Count;
                    }
                    var cells=new object[columnCount];
                    for (int i=0; i<columnCount; i++) {
                        cells[i]=i<fields.Count ? fields[i] : "";
                    }
                    table.Rows.Add(cells);
                }
            }
            // Drop trailing empty rows (all cells blank).
            for (int r=table.Rows.Count-1; r>=0; r--) {
                bool empty=true;
                foreach (object cell in table.Rows[r].ItemArray) {
                    if (cell!=null && !string.IsNullOrEmpty(cell.ToString().Trim())) {
                        empty=false;
                        break;
                    }
                }
                if (empty) {
                    table.Rows.RemoveAt(r);
                } else {
                    break;
                }
            }
            dataSet.Tables.Add(table);
            return dataSet;
        }

        internal static List<string> ParseLine(string line) {
            var fields=new List<string>();
            var sb=new StringBuilder();
            bool inQuotes=false;
            for (int i=0; i<line.Length; i++) {
                char c=line[i];
                if (inQuotes) {
                    if (c=='"') {
                        if (i+1<line.Length && line[i+1]=='"') {
                            sb.Append('"');
                            i++;
                        } else {
                            inQuotes=false;
                        }
                    } else {
                        sb.Append(c);
                    }
                } else {
                    if (c=='"') {
                        inQuotes=true;
                    } else if (c==',') {
                        fields.Add(sb.ToString());
                        sb.Length=0;
                    } else {
                        sb.Append(c);
                    }
                }
            }
            fields.Add(sb.ToString());
            return fields;
        }
    }
}
