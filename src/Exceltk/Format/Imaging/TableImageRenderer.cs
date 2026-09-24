using System;
using System.Collections.Generic;
using System.Text;

using Exceltk.Format;

namespace Exceltk.Format.Imaging {
    /// <summary>
    /// Renders a simple RGBA preview of a table. Precise data lives in the PNG marker chunk.
    /// </summary>
    public static class TableImageRenderer {
        private const int Pad = 8;
        private const int CellPadX = 6;
        private const int CellPadY = 4;
        private const int CharW = 6;
        private const int CharH = 8;

        public static byte[] RenderPng(TableDocument document) {
            SheetDocument sheet = document.Sheets.Count > 0
                ? document.Sheets[0]
                : new SheetDocument { Name = "Sheet1", Rows = new List<List<string>>() };

            int rows = Math.Max(1, sheet.Rows.Count);
            int cols = 1;
            foreach (List<string> row in sheet.Rows) {
                if (row != null) {
                    cols = Math.Max(cols, row.Count);
                }
            }

            var display = new string[rows, cols];
            var widths = new int[cols];
            for (int c = 0; c < cols; c++) {
                widths[c] = 3;
            }
            for (int r = 0; r < rows; r++) {
                List<string> row = r < sheet.Rows.Count ? sheet.Rows[r] : null;
                for (int c = 0; c < cols; c++) {
                    string text = row != null && c < row.Count ? (row[c] ?? "") : "";
                    text = Sanitize(text);
                    display[r, c] = text;
                    widths[c] = Math.Max(widths[c], Math.Min(40, Math.Max(1, text.Length)));
                }
            }

            var cellW = new int[cols];
            int tableW = 1;
            for (int c = 0; c < cols; c++) {
                cellW[c] = widths[c] * CharW + CellPadX * 2;
                tableW += cellW[c] + 1;
            }
            int cellH = CharH + CellPadY * 2;
            int tableH = 1 + rows * (cellH + 1);

            int width = tableW + Pad * 2;
            int height = tableH + Pad * 2 + CharH + 4;
            var rgba = new byte[width * height * 4];
            Fill(rgba, width, height, 245, 248, 250, 255);

            // Title bar with sheet name / marker hint
            string title = "ExcelTk · " + (string.IsNullOrEmpty(sheet.Name) ? "Sheet" : sheet.Name) + " · marked";
            DrawText(rgba, width, height, Pad, Pad, title, 40, 60, 80, 255);

            int originX = Pad;
            int originY = Pad + CharH + 4;
            FillRect(rgba, width, height, originX, originY, tableW, tableH, 30, 40, 50, 255);

            int y = originY + 1;
            for (int r = 0; r < rows; r++) {
                int x = originX + 1;
                for (int c = 0; c < cols; c++) {
                    byte bgR = r == 0 ? (byte)230 : (byte)255;
                    byte bgG = r == 0 ? (byte)236 : (byte)255;
                    byte bgB = r == 0 ? (byte)242 : (byte)255;
                    FillRect(rgba, width, height, x, y, cellW[c], cellH, bgR, bgG, bgB, 255);
                    string text = Truncate(display[r, c], widths[c]);
                    DrawText(rgba, width, height, x + CellPadX, y + CellPadY, text, 20, 28, 36, 255);
                    x += cellW[c] + 1;
                }
                y += cellH + 1;
            }

            return MarkedPng.EncodeRgba(width, height, rgba, document.ToMarkedBytes());
        }

        private static string Sanitize(string text) {
            if (string.IsNullOrEmpty(text)) {
                return "";
            }
            var sb = new StringBuilder(text.Length);
            foreach (char ch in text) {
                if (ch == '\r' || ch == '\n' || ch == '\t') {
                    sb.Append(' ');
                } else {
                    sb.Append(ch);
                }
            }
            return sb.ToString();
        }

        private static string Truncate(string text, int maxChars) {
            if (text.Length <= maxChars) {
                return text;
            }
            if (maxChars <= 1) {
                return text.Substring(0, maxChars);
            }
            return text.Substring(0, maxChars - 1) + "…";
        }

        private static void Fill(byte[] rgba, int width, int height, byte r, byte g, byte b, byte a) {
            for (int i = 0; i < width * height; i++) {
                int o = i * 4;
                rgba[o] = r;
                rgba[o + 1] = g;
                rgba[o + 2] = b;
                rgba[o + 3] = a;
            }
        }

        private static void FillRect(byte[] rgba, int width, int height, int x, int y, int w, int h, byte r, byte g, byte b, byte a) {
            for (int yy = y; yy < y + h; yy++) {
                if (yy < 0 || yy >= height) {
                    continue;
                }
                for (int xx = x; xx < x + w; xx++) {
                    if (xx < 0 || xx >= width) {
                        continue;
                    }
                    int o = (yy * width + xx) * 4;
                    rgba[o] = r;
                    rgba[o + 1] = g;
                    rgba[o + 2] = b;
                    rgba[o + 3] = a;
                }
            }
        }

        private static void DrawText(byte[] rgba, int width, int height, int x, int y, string text, byte r, byte g, byte b, byte a) {
            int cursor = x;
            foreach (char ch in text) {
                byte[] glyph = BitmapFont.GetGlyph(ch);
                for (int gy = 0; gy < CharH; gy++) {
                    for (int gx = 0; gx < 5; gx++) {
                        if (((glyph[gy] >> (4 - gx)) & 1) == 0) {
                            continue;
                        }
                        int px = cursor + gx;
                        int py = y + gy;
                        if (px < 0 || py < 0 || px >= width || py >= height) {
                            continue;
                        }
                        int o = (py * width + px) * 4;
                        rgba[o] = r;
                        rgba[o + 1] = g;
                        rgba[o + 2] = b;
                        rgba[o + 3] = a;
                    }
                }
                cursor += CharW;
            }
        }
    }

    /// <summary>Tiny 5x8 monospace glyphs for PNG preview text.</summary>
    internal static class BitmapFont {
        private static readonly Dictionary<char, byte[]> Glyphs = Build();

        public static byte[] GetGlyph(char ch) {
            if (ch >= 'a' && ch <= 'z') {
                ch = char.ToUpperInvariant(ch);
            }
            byte[] glyph;
            if (Glyphs.TryGetValue(ch, out glyph)) {
                return glyph;
            }
            return Glyphs['?'];
        }

        private static Dictionary<char, byte[]> Build() {
            var map = new Dictionary<char, byte[]>();
            void Add(char c, params byte[] rows) { map[c] = rows; }

            Add(' ', 0, 0, 0, 0, 0, 0, 0, 0);
            Add('.', 0, 0, 0, 0, 0, 0, 0x04, 0);
            Add(',', 0, 0, 0, 0, 0, 0x04, 0x08, 0);
            Add(':', 0, 0, 0x04, 0, 0x04, 0, 0, 0);
            Add('-', 0, 0, 0, 0x1F, 0, 0, 0, 0);
            Add('_', 0, 0, 0, 0, 0, 0, 0x1F, 0);
            Add('+', 0, 0x04, 0x04, 0x1F, 0x04, 0x04, 0, 0);
            Add('/', 0, 0x01, 0x02, 0x04, 0x08, 0x10, 0, 0);
            Add('\\', 0, 0x10, 0x08, 0x04, 0x02, 0x01, 0, 0);
            Add('(', 0x02, 0x04, 0x08, 0x08, 0x08, 0x04, 0x02, 0);
            Add(')', 0x08, 0x04, 0x02, 0x02, 0x02, 0x04, 0x08, 0);
            Add('[', 0x0E, 0x08, 0x08, 0x08, 0x08, 0x08, 0x0E, 0);
            Add(']', 0x0E, 0x02, 0x02, 0x02, 0x02, 0x02, 0x0E, 0);
            Add('\'', 0x04, 0x04, 0, 0, 0, 0, 0, 0);
            Add('"', 0x0A, 0x0A, 0, 0, 0, 0, 0, 0);
            Add('!', 0x04, 0x04, 0x04, 0x04, 0, 0x04, 0, 0);
            Add('?', 0x0E, 0x11, 0x01, 0x02, 0x04, 0, 0x04, 0);
            Add('·', 0, 0, 0, 0x04, 0, 0, 0, 0);
            Add('…', 0, 0, 0, 0, 0, 0x15, 0, 0);

            Add('0', 0x0E, 0x11, 0x13, 0x15, 0x19, 0x11, 0x0E, 0);
            Add('1', 0x04, 0x0C, 0x04, 0x04, 0x04, 0x04, 0x0E, 0);
            Add('2', 0x0E, 0x11, 0x01, 0x02, 0x04, 0x08, 0x1F, 0);
            Add('3', 0x1F, 0x02, 0x04, 0x02, 0x01, 0x11, 0x0E, 0);
            Add('4', 0x02, 0x06, 0x0A, 0x12, 0x1F, 0x02, 0x02, 0);
            Add('5', 0x1F, 0x10, 0x1E, 0x01, 0x01, 0x11, 0x0E, 0);
            Add('6', 0x06, 0x08, 0x10, 0x1E, 0x11, 0x11, 0x0E, 0);
            Add('7', 0x1F, 0x01, 0x02, 0x04, 0x08, 0x08, 0x08, 0);
            Add('8', 0x0E, 0x11, 0x11, 0x0E, 0x11, 0x11, 0x0E, 0);
            Add('9', 0x0E, 0x11, 0x11, 0x0F, 0x01, 0x02, 0x0C, 0);

            Add('A', 0x0E, 0x11, 0x11, 0x1F, 0x11, 0x11, 0x11, 0);
            Add('B', 0x1E, 0x11, 0x11, 0x1E, 0x11, 0x11, 0x1E, 0);
            Add('C', 0x0E, 0x11, 0x10, 0x10, 0x10, 0x11, 0x0E, 0);
            Add('D', 0x1E, 0x11, 0x11, 0x11, 0x11, 0x11, 0x1E, 0);
            Add('E', 0x1F, 0x10, 0x10, 0x1E, 0x10, 0x10, 0x1F, 0);
            Add('F', 0x1F, 0x10, 0x10, 0x1E, 0x10, 0x10, 0x10, 0);
            Add('G', 0x0E, 0x11, 0x10, 0x17, 0x11, 0x11, 0x0F, 0);
            Add('H', 0x11, 0x11, 0x11, 0x1F, 0x11, 0x11, 0x11, 0);
            Add('I', 0x0E, 0x04, 0x04, 0x04, 0x04, 0x04, 0x0E, 0);
            Add('J', 0x07, 0x02, 0x02, 0x02, 0x02, 0x12, 0x0C, 0);
            Add('K', 0x11, 0x12, 0x14, 0x18, 0x14, 0x12, 0x11, 0);
            Add('L', 0x10, 0x10, 0x10, 0x10, 0x10, 0x10, 0x1F, 0);
            Add('M', 0x11, 0x1B, 0x15, 0x15, 0x11, 0x11, 0x11, 0);
            Add('N', 0x11, 0x19, 0x15, 0x13, 0x11, 0x11, 0x11, 0);
            Add('O', 0x0E, 0x11, 0x11, 0x11, 0x11, 0x11, 0x0E, 0);
            Add('P', 0x1E, 0x11, 0x11, 0x1E, 0x10, 0x10, 0x10, 0);
            Add('Q', 0x0E, 0x11, 0x11, 0x11, 0x15, 0x12, 0x0D, 0);
            Add('R', 0x1E, 0x11, 0x11, 0x1E, 0x14, 0x12, 0x11, 0);
            Add('S', 0x0F, 0x10, 0x10, 0x0E, 0x01, 0x01, 0x1E, 0);
            Add('T', 0x1F, 0x04, 0x04, 0x04, 0x04, 0x04, 0x04, 0);
            Add('U', 0x11, 0x11, 0x11, 0x11, 0x11, 0x11, 0x0E, 0);
            Add('V', 0x11, 0x11, 0x11, 0x11, 0x11, 0x0A, 0x04, 0);
            Add('W', 0x11, 0x11, 0x11, 0x15, 0x15, 0x1B, 0x11, 0);
            Add('X', 0x11, 0x11, 0x0A, 0x04, 0x0A, 0x11, 0x11, 0);
            Add('Y', 0x11, 0x11, 0x0A, 0x04, 0x04, 0x04, 0x04, 0);
            Add('Z', 0x1F, 0x01, 0x02, 0x04, 0x08, 0x10, 0x1F, 0);

            return map;
        }
    }
}
