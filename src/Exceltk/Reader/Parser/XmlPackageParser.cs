using System.Collections.Generic;
using System.Xml;
using Exceltk.Reader.Package;
using Exceltk.Reader.Xml;

namespace Exceltk.Reader.Parser {
    /// <summary>
    /// Streams <see cref="XmlPackage"/> entities from an OpenXML worksheet XmlReader.
    /// </summary>
    internal sealed class XmlPackageParser : IPackageParser<XmlPackage> {
        private readonly XmlReader m_reader;
        private readonly string m_namespaceUri;

        public XmlPackageParser(XmlReader reader, string namespaceUri = null) {
            m_reader = reader;
            m_namespaceUri = namespaceUri;
        }

        /// <summary>
        /// Walk the worksheet document, yielding dimension / row / sheetData-end packages.
        /// Caller should position at the start of the worksheet document.
        /// </summary>
        public IEnumerable<XmlPackage> Parse() {
            while (m_reader.Read()) {
                if (m_reader.NodeType != XmlNodeType.Element) {
                    continue;
                }

                if (m_reader.LocalName == XlsxWorksheet.N_worksheet) {
                    yield return new XmlWorksheetStartPackage {
                        NamespaceUri = m_reader.NamespaceURI
                    };
                } else if (m_reader.LocalName == XlsxWorksheet.N_dimension) {
                    yield return new XmlDimensionPackage {
                        Ref = m_reader.GetAttribute(XlsxWorksheet.A_ref)
                    };
                } else if (m_reader.LocalName == XlsxWorksheet.N_sheetData) {
                    if (m_reader.IsEmptyElement) {
                        yield return new XmlSheetDataEndPackage();
                        continue;
                    }
                    foreach (XmlPackage rowPackage in ParseRows()) {
                        yield return rowPackage;
                    }
                    yield return new XmlSheetDataEndPackage();
                } else if (m_reader.LocalName == XlsxWorksheet.N_mergeCells) {
                    foreach (XmlMergePackage merge in ParseMergesFromCurrent()) {
                        yield return merge;
                    }
                } else if (m_reader.LocalName == XlsxWorksheet.N_hyperlinks) {
                    foreach (XmlHyperlinkPackage link in ParseHyperlinksFromCurrent()) {
                        yield return link;
                    }
                }
            }
        }

        /// <summary>
        /// Stream row packages from the current sheetData element (or subsequent rows).
        /// </summary>
        public IEnumerable<XmlRowPackage> ParseRows() {
            bool first = true;
            while (true) {
                bool isRow;
                if (first && m_reader.NodeType == XmlNodeType.Element &&
                    m_reader.LocalName == XlsxWorksheet.N_sheetData) {
                    if (m_namespaceUri != null) {
                        isRow = m_reader.ReadToFollowing(XlsxWorksheet.N_row, m_namespaceUri);
                    } else {
                        isRow = m_reader.ReadToFollowing(XlsxWorksheet.N_row);
                    }
                    first = false;
                } else {
                    if (m_reader.LocalName == XlsxWorksheet.N_row && m_reader.NodeType == XmlNodeType.EndElement) {
                        m_reader.Read();
                    }
                    isRow = (m_reader.NodeType == XmlNodeType.Element && m_reader.LocalName == XlsxWorksheet.N_row);
                    first = false;
                }

                if (!isRow) {
                    yield break;
                }

                yield return ReadCurrentRow();
            }
        }

        public IEnumerable<XmlMergePackage> ParseMerges() {
            if (!m_reader.ReadToFollowing(XlsxWorksheet.N_mergeCells)) {
                yield break;
            }
            foreach (XmlMergePackage merge in ParseMergesFromCurrent()) {
                yield return merge;
            }
        }

        public IEnumerable<XmlHyperlinkPackage> ParseHyperlinks() {
            if (!m_reader.ReadToFollowing(XlsxWorksheet.N_hyperlinks)) {
                yield break;
            }
            foreach (XmlHyperlinkPackage link in ParseHyperlinksFromCurrent()) {
                yield return link;
            }
        }

        private IEnumerable<XmlMergePackage> ParseMergesFromCurrent() {
            if (m_reader.IsEmptyElement) {
                yield break;
            }

            while (m_reader.Read()) {
                if (m_reader.NodeType != XmlNodeType.Element) {
                    if (m_reader.NodeType == XmlNodeType.EndElement &&
                        m_reader.LocalName == XlsxWorksheet.N_mergeCells) {
                        yield break;
                    }
                    continue;
                }
                if (m_reader.LocalName != XlsxWorksheet.N_mergeCell) {
                    yield break;
                }
                yield return new XmlMergePackage {
                    Ref = m_reader.GetAttribute(XlsxWorksheet.A_ref)
                };
            }
        }

        private IEnumerable<XmlHyperlinkPackage> ParseHyperlinksFromCurrent() {
            if (m_reader.IsEmptyElement) {
                yield break;
            }

            while (m_reader.Read()) {
                if (m_reader.NodeType != XmlNodeType.Element) {
                    yield break;
                }
                if (m_reader.LocalName != XlsxWorksheet.N_hyperlink) {
                    yield break;
                }
                yield return new XmlHyperlinkPackage {
                    Ref = m_reader.GetAttribute(XlsxWorksheet.A_ref),
                    Display = m_reader.GetAttribute(XlsxWorksheet.A_display),
                    RelationshipId = m_reader.GetAttribute(XlsxWorksheet.A_rid),
                    Location = m_reader.GetAttribute("location")
                };
            }
        }

        private XmlRowPackage ReadCurrentRow() {
            var row = new XmlRowPackage {
                RowIndexAttribute = m_reader.GetAttribute(XlsxWorksheet.A_r)
            };

            string a_s = string.Empty;
            string a_t = string.Empty;
            string a_r = string.Empty;
            string formula = null;
            bool hasValue = false;
            bool hasFormula = false;

            while (m_reader.Read()) {
                if (m_reader.Depth == 2) {
                    break;
                }

                if (m_reader.NodeType == XmlNodeType.Element) {
                    hasValue = false;
                    if (m_reader.LocalName == XlsxWorksheet.N_c) {
                        a_s = m_reader.GetAttribute(XlsxWorksheet.A_s);
                        a_t = m_reader.GetAttribute(XlsxWorksheet.A_t);
                        a_r = m_reader.GetAttribute(XlsxWorksheet.A_r);
                        formula = null;
                    } else if (m_reader.LocalName == XlsxWorksheet.N_f) {
                        hasFormula = true;
                    } else if (m_reader.LocalName == XlsxWorksheet.N_v || m_reader.LocalName == XlsxWorksheet.N_t) {
                        hasValue = true;
                        hasFormula = false;
                    }
                }

                if (m_reader.NodeType == XmlNodeType.Text && hasFormula) {
                    formula = m_reader.Value;
                    hasFormula = false;
                }

                if (m_reader.NodeType == XmlNodeType.Text && hasValue) {
                    row.Cells.Add(new XmlCellPackage {
                        Reference = a_r,
                        StyleId = a_s,
                        CellType = a_t,
                        Formula = formula,
                        ValueText = m_reader.Value
                    });
                    formula = null;
                    hasValue = false;
                }
            }

            return row;
        }
    }
}
