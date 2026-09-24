using System;
using System.IO;
using System.Text;

using ICSharpCode.SharpZipLib.Checksums;
using ICSharpCode.SharpZipLib.Zip.Compression;

namespace Exceltk.Format.Imaging {
    /// <summary>
    /// Minimal PNG codec that can embed/extract an ExcelTk private chunk (<c>tkXl</c>).
    /// Unmarked PNGs are rejected on import.
    /// </summary>
    public static class MarkedPng {
        public const string ChunkType = "tkXl";
        private static readonly byte[] Signature = { 137, 80, 78, 71, 13, 10, 26, 10 };

        public static byte[] EncodeRgba(int width, int height, byte[] rgba, byte[] markerPayload) {
            if (width <= 0 || height <= 0) {
                throw new ArgumentOutOfRangeException("width/height");
            }
            if (rgba == null || rgba.Length < width * height * 4) {
                throw new ArgumentException("RGBA buffer too small.");
            }
            if (markerPayload == null || markerPayload.Length == 0) {
                throw new ArgumentException("ExcelTk marker payload is required.");
            }

            byte[] raw = BuildRawImage(width, height, rgba);
            byte[] compressed = Deflate(raw);

            using (var ms = new MemoryStream()) {
                ms.Write(Signature, 0, Signature.Length);
                WriteChunk(ms, "IHDR", BuildIhdr(width, height));
                WriteChunk(ms, ChunkType, markerPayload);
                WriteChunk(ms, "IDAT", compressed);
                WriteChunk(ms, "IEND", Array.Empty<byte>());
                return ms.ToArray();
            }
        }

        public static bool TryReadMarker(byte[] pngBytes, out byte[] markerPayload) {
            markerPayload = null;
            if (pngBytes == null || pngBytes.Length < 8) {
                return false;
            }
            for (int i = 0; i < 8; i++) {
                if (pngBytes[i] != Signature[i]) {
                    return false;
                }
            }

            int offset = 8;
            while (offset + 12 <= pngBytes.Length) {
                int length = ReadInt32(pngBytes, offset);
                if (length < 0 || offset + 12 + length > pngBytes.Length) {
                    return false;
                }
                string type = Encoding.ASCII.GetString(pngBytes, offset + 4, 4);
                if (type == ChunkType) {
                    markerPayload = new byte[length];
                    Buffer.BlockCopy(pngBytes, offset + 8, markerPayload, 0, length);
                    return true;
                }
                if (type == "IEND") {
                    break;
                }
                offset += 12 + length;
            }
            return false;
        }

        public static byte[] RequireMarker(byte[] pngBytes) {
            byte[] payload;
            if (!TryReadMarker(pngBytes, out payload)) {
                throw new FormatException(
                    "Image is not an ExcelTk marked PNG (missing '" + ChunkType + "' chunk). Unmarked images are not supported.");
            }
            return payload;
        }

        private static byte[] BuildRawImage(int width, int height, byte[] rgba) {
            var raw = new byte[(width * 4 + 1) * height];
            int dst = 0;
            int src = 0;
            for (int y = 0; y < height; y++) {
                raw[dst++] = 0; // filter None
                int rowBytes = width * 4;
                Buffer.BlockCopy(rgba, src, raw, dst, rowBytes);
                dst += rowBytes;
                src += rowBytes;
            }
            return raw;
        }

        private static byte[] BuildIhdr(int width, int height) {
            var data = new byte[13];
            WriteInt32(data, 0, width);
            WriteInt32(data, 4, height);
            data[8] = 8;  // bit depth
            data[9] = 6;  // RGBA
            data[10] = 0;
            data[11] = 0;
            data[12] = 0;
            return data;
        }

        private static byte[] Deflate(byte[] input) {
            var deflater = new Deflater(Deflater.DEFAULT_COMPRESSION, false);
            using (var output = new MemoryStream()) {
                var buffer = new byte[1024];
                deflater.SetInput(input);
                deflater.Finish();
                while (!deflater.IsFinished) {
                    int count = deflater.Deflate(buffer);
                    if (count <= 0) {
                        break;
                    }
                    output.Write(buffer, 0, count);
                }
                return output.ToArray();
            }
        }

        private static void WriteChunk(Stream stream, string type, byte[] data) {
            byte[] typeBytes = Encoding.ASCII.GetBytes(type);
            if (typeBytes.Length != 4) {
                throw new ArgumentException("PNG chunk type must be 4 bytes: " + type);
            }
            WriteInt32(stream, data.Length);
            stream.Write(typeBytes, 0, 4);
            if (data.Length > 0) {
                stream.Write(data, 0, data.Length);
            }

            var crc = new Crc32();
            crc.Update(typeBytes);
            if (data.Length > 0) {
                crc.Update(data);
            }
            WriteInt32(stream, (int)crc.Value);
        }

        private static void WriteInt32(Stream stream, int value) {
            stream.WriteByte((byte)((value >> 24) & 0xff));
            stream.WriteByte((byte)((value >> 16) & 0xff));
            stream.WriteByte((byte)((value >> 8) & 0xff));
            stream.WriteByte((byte)(value & 0xff));
        }

        private static void WriteInt32(byte[] buffer, int offset, int value) {
            buffer[offset] = (byte)((value >> 24) & 0xff);
            buffer[offset + 1] = (byte)((value >> 16) & 0xff);
            buffer[offset + 2] = (byte)((value >> 8) & 0xff);
            buffer[offset + 3] = (byte)(value & 0xff);
        }

        private static int ReadInt32(byte[] buffer, int offset) {
            return (buffer[offset] << 24)
                | (buffer[offset + 1] << 16)
                | (buffer[offset + 2] << 8)
                | buffer[offset + 3];
        }
    }
}
