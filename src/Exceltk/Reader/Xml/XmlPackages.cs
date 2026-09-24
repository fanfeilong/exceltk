using System.Collections.Generic;
using System.Xml;
using Exceltk.Reader.Package;

namespace Exceltk.Reader.Xml {
    /// <summary>SpreadsheetML <c>dimension</c> element (<c>ref</c> attribute).</summary>
    internal sealed class XmlDimensionPackage : XmlPackage {
        public string Ref {
            get;
            private set;
        }

        public override void Decode(XmlReader reader) {
            Ref = reader.GetAttribute(XlsxWorksheet.A_ref);
            if (!reader.IsEmptyElement) {
                reader.Skip();
            }
        }

        public override void EncodeBody(XmlWriter writer) {
            writer.WriteStartElement(XlsxWorksheet.N_dimension);
            if (Ref != null) {
                writer.WriteAttributeString(XlsxWorksheet.A_ref, Ref);
            }
            writer.WriteEndElement();
        }
    }

    /// <summary>
    /// SpreadsheetML <c>c</c> (cell) element: attributes <c>r</c>/<c>t</c>/<c>s</c>,
    /// children <c>v</c>/<c>t</c>/<c>f</c> as Office defines.
    /// </summary>
    internal sealed class XmlCellPackage : XmlPackage {
        public string Reference {
            get;
            private set;
        }

        public string StyleId {
            get;
            private set;
        }

        public string CellType {
            get;
            private set;
        }

        public string Formula {
            get;
            private set;
        }

        public string ValueText {
            get;
            private set;
        }

        public override void Decode(XmlReader reader) {
            Reference = reader.GetAttribute(XlsxWorksheet.A_r);
            StyleId = reader.GetAttribute(XlsxWorksheet.A_s);
            CellType = reader.GetAttribute(XlsxWorksheet.A_t);
            Formula = null;
            ValueText = null;

            if (reader.IsEmptyElement) {
                return;
            }

            bool hasValue = false;
            bool hasFormula = false;

            while (reader.Read()) {
                if (reader.NodeType == XmlNodeType.EndElement && reader.LocalName == XlsxWorksheet.N_c) {
                    break;
                }

                if (reader.NodeType == XmlNodeType.Element) {
                    hasValue = false;
                    if (reader.LocalName == XlsxWorksheet.N_f) {
                        hasFormula = true;
                    } else if (reader.LocalName == XlsxWorksheet.N_v || reader.LocalName == XlsxWorksheet.N_t) {
                        hasValue = true;
                        hasFormula = false;
                    }
                }

                if (reader.NodeType == XmlNodeType.Text && hasFormula) {
                    Formula = reader.Value;
                    hasFormula = false;
                }

                if (reader.NodeType == XmlNodeType.Text && hasValue) {
                    ValueText = reader.Value;
                    hasValue = false;
                }
            }
        }

        public override void EncodeBody(XmlWriter writer) {
            writer.WriteStartElement(XlsxWorksheet.N_c);
            if (Reference != null) {
                writer.WriteAttributeString(XlsxWorksheet.A_r, Reference);
            }
            if (StyleId != null) {
                writer.WriteAttributeString(XlsxWorksheet.A_s, StyleId);
            }
            if (CellType != null) {
                writer.WriteAttributeString(XlsxWorksheet.A_t, CellType);
            }
            if (Formula != null) {
                writer.WriteElementString(XlsxWorksheet.N_f, Formula);
            }
            if (ValueText != null) {
                string valueLocal = CellType == "inlineStr" ? XlsxWorksheet.N_t : XlsxWorksheet.N_v;
                writer.WriteElementString(valueLocal, ValueText);
            }
            writer.WriteEndElement();
        }
    }

    /// <summary>
    /// SpreadsheetML <c>row</c> element: row attributes plus child <c>c</c> cell packages.
    /// </summary>
    internal sealed class XmlRowPackage : XmlPackage {
        public XmlRowPackage() {
            Cells = new List<XmlCellPackage>();
        }

        public string RowIndexAttribute {
            get;
            private set;
        }

        public List<XmlCellPackage> Cells {
            get;
            private set;
        }

        public override void Decode(XmlReader reader) {
            RowIndexAttribute = reader.GetAttribute(XlsxWorksheet.A_r);
            Cells.Clear();

            if (reader.IsEmptyElement) {
                return;
            }

            while (reader.Read()) {
                if (reader.NodeType == XmlNodeType.EndElement && reader.LocalName == XlsxWorksheet.N_row) {
                    break;
                }

                if (reader.NodeType == XmlNodeType.Element && reader.LocalName == XlsxWorksheet.N_c) {
                    var cell = new XmlCellPackage();
                    cell.Decode(reader);
                    // Match prior sheet reader: only cells that carried a value text node.
                    if (cell.ValueText != null) {
                        Cells.Add(cell);
                    }
                }
            }
        }

        public override void EncodeBody(XmlWriter writer) {
            writer.WriteStartElement(XlsxWorksheet.N_row);
            if (RowIndexAttribute != null) {
                writer.WriteAttributeString(XlsxWorksheet.A_r, RowIndexAttribute);
            }
            foreach (XmlCellPackage cell in Cells) {
                cell.EncodeBody(writer);
            }
            writer.WriteEndElement();
        }
    }

    /// <summary>SpreadsheetML <c>mergeCell</c> element (<c>ref</c> attribute).</summary>
    internal sealed class XmlMergePackage : XmlPackage {
        public string Ref {
            get;
            private set;
        }

        public override void Decode(XmlReader reader) {
            Ref = reader.GetAttribute(XlsxWorksheet.A_ref);
            if (!reader.IsEmptyElement) {
                reader.Skip();
            }
        }

        public override void EncodeBody(XmlWriter writer) {
            writer.WriteStartElement(XlsxWorksheet.N_mergeCell);
            if (Ref != null) {
                writer.WriteAttributeString(XlsxWorksheet.A_ref, Ref);
            }
            writer.WriteEndElement();
        }
    }

    /// <summary>SpreadsheetML <c>hyperlink</c> element.</summary>
    internal sealed class XmlHyperlinkPackage : XmlPackage {
        public string Ref {
            get;
            private set;
        }

        public string Display {
            get;
            private set;
        }

        public string RelationshipId {
            get;
            private set;
        }

        public string Location {
            get;
            private set;
        }

        public override void Decode(XmlReader reader) {
            Ref = reader.GetAttribute(XlsxWorksheet.A_ref);
            Display = reader.GetAttribute(XlsxWorksheet.A_display);
            RelationshipId = reader.GetAttribute(XlsxWorksheet.A_rid);
            Location = reader.GetAttribute("location");
            if (!reader.IsEmptyElement) {
                reader.Skip();
            }
        }

        public override void EncodeBody(XmlWriter writer) {
            writer.WriteStartElement(XlsxWorksheet.N_hyperlink);
            if (Ref != null) {
                writer.WriteAttributeString(XlsxWorksheet.A_ref, Ref);
            }
            if (Display != null) {
                writer.WriteAttributeString(XlsxWorksheet.A_display, Display);
            }
            if (RelationshipId != null) {
                writer.WriteAttributeString(XlsxWorksheet.A_rid, RelationshipId);
            }
            if (Location != null) {
                writer.WriteAttributeString("location", Location);
            }
            writer.WriteEndElement();
        }
    }
}
