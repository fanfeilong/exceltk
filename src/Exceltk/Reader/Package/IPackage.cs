using System;
using System.IO;
using System.Xml;
using Exceltk.Reader.Xml;

namespace Exceltk.Reader.Package {
    /// <summary>
    /// Office-defined protocol unit (BIFF record / SpreadsheetML element).
    /// Owns Decode of Office-defined fields — not a UI progressive-render DTO.
    /// </summary>
    public interface IPackage {
        /// <summary>
        /// Decode Office-defined body bytes into this package's fields.
        /// Called by <c>PackageParser</c> after the body is fully accumulated.
        /// </summary>
        void DecodeBody(byte[] data, int offset, int count);
    }

    /// <summary>
    /// SpreadsheetML element package base. Concrete units live in <c>Reader.Xml</c>.
    /// Field layout is Office-defined; each subclass owns <see cref="Decode"/> / <see cref="EncodeBody"/>.
    /// Binary BIFF packages live in <c>Reader.Binary</c> as <c>BinaryPackage</c> / <c>XlsBiff*</c>.
    /// </summary>
    public abstract class XmlPackage : IPackage {
        /// <summary>
        /// Decode from an <see cref="XmlReader"/> positioned on this element's start tag.
        /// Consumes the element (including children) per SpreadsheetML.
        /// </summary>
        public abstract void Decode(XmlReader reader);

        /// <summary>
        /// Encode this package as SpreadsheetML into <paramref name="writer"/>.
        /// </summary>
        public abstract void EncodeBody(XmlWriter writer);

        /// <summary>
        /// Decode a complete element XML fragment (UTF-8). Used when the framer owns bytes.
        /// </summary>
        public virtual void DecodeBody(byte[] data, int offset, int count) {
            if (data == null || count <= 0) {
                return;
            }

            var settings = new XmlReaderSettings {
                ConformanceLevel = ConformanceLevel.Fragment,
                IgnoreComments = true,
                IgnoreWhitespace = true,
                CloseInput = true
            };

            using (var stream = new MemoryStream(data, offset, count, writable: false))
            using (XmlReader reader = XmlReader.Create(stream, settings)) {
                if (reader.Read() && reader.NodeType == XmlNodeType.Element) {
                    Decode(reader);
                }
            }
        }

        /// <summary>
        /// Encode this package to a UTF-8 SpreadsheetML fragment.
        /// Returns bytes written, or -1 if the buffer is too small.
        /// </summary>
        public virtual int EncodeBody(byte[] buffer, int offset, int capacity) {
            using (var ms = new MemoryStream()) {
                var settings = new XmlWriterSettings {
                    Encoding = new System.Text.UTF8Encoding(false),
                    OmitXmlDeclaration = true,
                    ConformanceLevel = ConformanceLevel.Fragment
                };
                using (XmlWriter writer = XmlWriter.Create(ms, settings)) {
                    EncodeBody(writer);
                }
                byte[] bytes = ms.ToArray();
                if (capacity < bytes.Length) {
                    return -1;
                }
                Buffer.BlockCopy(bytes, 0, buffer, offset, bytes.Length);
                return bytes.Length;
            }
        }

        /// <summary>Create an empty package shell for a SpreadsheetML local name, or null if not a package unit.</summary>
        public static XmlPackage CreateEmpty(string localName) {
            if (localName == XlsxWorksheet.N_dimension) {
                return new XmlDimensionPackage();
            }
            if (localName == XlsxWorksheet.N_row) {
                return new XmlRowPackage();
            }
            if (localName == XlsxWorksheet.N_c) {
                return new XmlCellPackage();
            }
            if (localName == XlsxWorksheet.N_mergeCell) {
                return new XmlMergePackage();
            }
            if (localName == XlsxWorksheet.N_hyperlink) {
                return new XmlHyperlinkPackage();
            }
            return null;
        }
    }
}
