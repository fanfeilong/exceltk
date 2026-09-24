using System;

namespace Exceltk.Format {
    /// <summary>
    /// One exported sheet/file produced by a format plugin.
    /// </summary>
    public sealed class FormatArtifact {
        public string Name { get; set; }
        public string Extension { get; set; }
        public bool IsBinary { get; set; }
        public string TextContent { get; set; }
        public byte[] BinaryContent { get; set; }

        public static FormatArtifact Text(string name, string extension, string content) {
            return new FormatArtifact {
                Name = name ?? "",
                Extension = NormalizeExtension(extension),
                IsBinary = false,
                TextContent = content ?? ""
            };
        }

        public static FormatArtifact Binary(string name, string extension, byte[] content) {
            return new FormatArtifact {
                Name = name ?? "",
                Extension = NormalizeExtension(extension),
                IsBinary = true,
                BinaryContent = content ?? Array.Empty<byte>()
            };
        }

        private static string NormalizeExtension(string extension) {
            if (string.IsNullOrEmpty(extension)) {
                return "";
            }
            return extension.StartsWith(".", StringComparison.Ordinal)
                ? extension.Substring(1)
                : extension;
        }
    }
}
