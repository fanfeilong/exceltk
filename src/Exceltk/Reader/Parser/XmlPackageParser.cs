using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using Exceltk.Reader.Package;
using Exceltk.Reader.Xml;

namespace Exceltk.Reader.Parser {
    /// <summary>
    /// SpreadsheetML element framer (XUdt-style).
    /// Recognizes Office elements (<c>dimension</c>/<c>row</c>/<c>c</c>/<c>mergeCell</c>/<c>hyperlink</c>),
    /// creates empty packages, calls <see cref="XmlPackage.Decode"/>, and queues them.
    /// Container elements (<c>worksheet</c>/<c>sheetData</c>/…) are parser state only — not packages.
    /// </summary>
    internal sealed class XmlPackageParser : IPackageParser<XmlPackage> {
        private readonly Queue<XmlPackage> m_queue = new Queue<XmlPackage>();
        private readonly MemoryStream m_pushBuffer = new MemoryStream();

        private XmlReader m_reader;
        private string m_namespaceUri;
        private bool m_ownsReader;
        private bool m_eof;
        private string m_stopAtEndElement;
        private bool m_inSheetData;
        private bool m_pushMode;

        /// <summary>
        /// Frame from an existing <see cref="XmlReader"/> (file or network stream).
        /// Call <see cref="Pump"/> / <see cref="TryFrameOne"/> to fill the package queue.
        /// </summary>
        public XmlPackageParser(XmlReader reader, string namespaceUri = null) {
            if (reader == null) {
                throw new ArgumentNullException("reader");
            }
            m_reader = reader;
            m_namespaceUri = namespaceUri;
            m_ownsReader = false;
            m_pushMode = false;

            if (m_reader.NodeType == XmlNodeType.Element &&
                m_reader.LocalName == XlsxWorksheet.N_sheetData) {
                m_inSheetData = !m_reader.IsEmptyElement;
            }
        }

        /// <summary>
        /// PushData-first mode: accumulate UTF-8 worksheet bytes, then frame via an internal reader.
        /// </summary>
        public XmlPackageParser(string namespaceUri = null) {
            m_namespaceUri = namespaceUri;
            m_pushMode = true;
            m_ownsReader = true;
        }

        /// <summary>Namespace URI captured from the <c>worksheet</c> element (parser state, not a package).</summary>
        public string NamespaceUri {
            get {
                return m_namespaceUri;
            }
        }

        public bool HavePackage {
            get {
                return m_queue.Count > 0;
            }
        }

        public bool HaveUnParsedData {
            get {
                if (m_pushMode && m_pushBuffer.Length > 0 && m_reader == null) {
                    return true;
                }
                return !m_eof && m_reader != null;
            }
        }

        public bool IsEof {
            get {
                return m_eof;
            }
        }

        public void Reset() {
            m_queue.Clear();
            m_eof = false;
            m_stopAtEndElement = null;
            m_inSheetData = false;
            if (m_ownsReader && m_reader != null) {
                m_reader.Close();
                m_reader = null;
            }
            m_pushBuffer.SetLength(0);
            m_pushBuffer.Position = 0;
        }

        public XmlPackage PopPackage() {
            return m_queue.Dequeue();
        }

        /// <summary>
        /// Feed UTF-8 SpreadsheetML bytes (push mode). In XmlReader mode, bytes are ignored
        /// because the reader already owns the input stream (e.g. NetworkStream).
        /// </summary>
        public void PushData(byte[] data, int offset, int count) {
            if (data == null) {
                throw new ArgumentNullException("data");
            }
            if (offset < 0 || count < 0 || offset + count > data.Length) {
                throw new ArgumentOutOfRangeException("count");
            }
            if (!m_pushMode) {
                return;
            }
            if (m_reader != null) {
                throw new InvalidOperationException("PushData after framing started; Reset first.");
            }
            m_pushBuffer.Write(data, offset, count);
        }

        /// <summary>
        /// Finish the push buffer and open an XmlReader over the accumulated worksheet XML.
        /// </summary>
        public void CompletePush() {
            if (!m_pushMode) {
                return;
            }
            EnsurePushReader();
        }

        /// <summary>
        /// Enter <c>sheetData</c> row framing: stop when <c>&lt;/sheetData&gt;</c> is reached.
        /// Caller should position on the non-empty <c>sheetData</c> start element.
        /// </summary>
        public void BeginSheetDataRows() {
            m_stopAtEndElement = XlsxWorksheet.N_sheetData;
            if (m_reader != null &&
                m_reader.NodeType == XmlNodeType.Element &&
                m_reader.LocalName == XlsxWorksheet.N_sheetData) {
                if (m_reader.IsEmptyElement) {
                    m_eof = true;
                    m_inSheetData = false;
                    return;
                }
                m_inSheetData = true;
            }
        }

        /// <summary>
        /// Advance to a container element (<c>mergeCells</c> / <c>hyperlinks</c> / …).
        /// Returns false if missing or empty. Framing stops at the matching end element.
        /// </summary>
        public bool SeekToElement(string localName) {
            if (m_reader == null) {
                return false;
            }

            bool found = m_reader.NodeType == XmlNodeType.Element && m_reader.LocalName == localName;
            if (!found) {
                if (m_namespaceUri != null) {
                    found = m_reader.ReadToFollowing(localName, m_namespaceUri);
                } else {
                    found = m_reader.ReadToFollowing(localName);
                }
            }

            if (!found) {
                return false;
            }
            if (m_reader.IsEmptyElement) {
                return false;
            }

            m_stopAtEndElement = localName;
            m_eof = false;
            return true;
        }

        /// <summary>
        /// Frame until at least one package is queued, or input is exhausted.
        /// </summary>
        public bool Pump() {
            EnsurePushReader();
            while (!HavePackage && !m_eof) {
                if (!TryFrameOne()) {
                    break;
                }
            }
            return HavePackage;
        }

        /// <summary>
        /// Advance the reader until one Office element is decoded and queued, or EOF / region end.
        /// </summary>
        public bool TryFrameOne() {
            EnsurePushReader();
            if (m_reader == null || m_eof) {
                return false;
            }

            while (m_reader.Read()) {
                if (m_reader.NodeType == XmlNodeType.EndElement) {
                    if (m_reader.LocalName == XlsxWorksheet.N_sheetData) {
                        m_inSheetData = false;
                    }
                    if (m_stopAtEndElement != null && m_reader.LocalName == m_stopAtEndElement) {
                        m_eof = true;
                        return false;
                    }
                    continue;
                }

                if (m_reader.NodeType != XmlNodeType.Element) {
                    continue;
                }

                string localName = m_reader.LocalName;

                if (localName == XlsxWorksheet.N_worksheet) {
                    if (m_namespaceUri == null) {
                        m_namespaceUri = m_reader.NamespaceURI;
                    }
                    continue;
                }

                if (localName == XlsxWorksheet.N_sheetData) {
                    m_inSheetData = !m_reader.IsEmptyElement;
                    if (!m_inSheetData && m_stopAtEndElement == XlsxWorksheet.N_sheetData) {
                        m_eof = true;
                        return false;
                    }
                    continue;
                }

                // Container wrappers — not packages.
                if (localName == XlsxWorksheet.N_mergeCells ||
                    localName == XlsxWorksheet.N_hyperlinks) {
                    continue;
                }

                XmlPackage package = XmlPackage.CreateEmpty(localName);
                if (package == null) {
                    continue;
                }

                package.Decode(m_reader);
                m_queue.Enqueue(package);
                return true;
            }

            m_eof = true;
            return false;
        }

        private void EnsurePushReader() {
            if (!m_pushMode || m_reader != null) {
                return;
            }
            if (m_pushBuffer.Length == 0) {
                return;
            }
            m_pushBuffer.Position = 0;
            var settings = new XmlReaderSettings {
                CloseInput = false,
                IgnoreComments = true,
                IgnoreWhitespace = false,
                ConformanceLevel = ConformanceLevel.Document
            };
            m_reader = XmlReader.Create(m_pushBuffer, settings);
            m_eof = false;
        }
    }
}
