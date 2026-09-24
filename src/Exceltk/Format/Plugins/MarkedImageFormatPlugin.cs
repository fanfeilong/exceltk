using System;
using System.Collections.Generic;
using System.IO;

using Exceltk.Format.Imaging;
using Exceltk.Reader;

namespace Exceltk.Format.Plugins {
    /// <summary>
    /// Export Excel sheets to PNG images that embed an ExcelTk marker chunk (`tkXl`).
    /// Import only accepts marked PNGs so round-trip conversion stays precise.
    /// </summary>
    public sealed class MarkedImageFormatPlugin : IFormatPlugin {
        public string Name { get { return "img"; } }
        public string FileExtension { get { return "png"; } }
        public string Description { get { return "Marked PNG image (export/import; unmarked images rejected)"; } }
        public bool SupportsExport { get { return true; } }
        public bool SupportsImport { get { return true; } }

        public IEnumerable<FormatArtifact> Export(DataSet dataSet, string sheetFilter) {
            DataSet shrinked = new DataSet();
            foreach (DataTable table in MarkdownFormatPlugin.FilterSheets(dataSet, sheetFilter)) {
                DataTable snapshot = MarkdownFormatPlugin.Snapshot(table);
                snapshot.Shrink();
                shrinked.Tables.Add(snapshot);
            }
            TableDocument full = TableDocument.FromDataSet(shrinked);
            foreach (SheetDocument sheet in full.Sheets) {
                var one = new TableDocument();
                one.Sheets.Add(sheet);
                byte[] png = TableImageRenderer.RenderPng(one);
                yield return FormatArtifact.Binary(sheet.Name, FileExtension, png);
            }
        }

        public DataSet Import(Stream input, string sourcePath) {
            byte[] bytes = ReadAll(input);
            byte[] marker = MarkedPng.RequireMarker(bytes);
            TableDocument document = TableDocument.FromMarkedBytes(marker);
            return document.ToDataSet();
        }

        private static byte[] ReadAll(Stream input) {
            if (input is MemoryStream ms && ms.TryGetBuffer(out ArraySegment<byte> segment)) {
                var copy = new byte[segment.Count];
                Buffer.BlockCopy(segment.Array, segment.Offset, copy, 0, segment.Count);
                return copy;
            }
            using (var buffer = new MemoryStream()) {
                input.CopyTo(buffer);
                return buffer.ToArray();
            }
        }
    }
}
