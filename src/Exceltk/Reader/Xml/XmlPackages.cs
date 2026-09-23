using System.Collections.Generic;
using Exceltk.Reader.Package;

namespace Exceltk.Reader.Xml {
    /// <summary>Emitted when the worksheet root element is seen.</summary>
    internal sealed class XmlWorksheetStartPackage : XmlPackage {
        public string NamespaceUri {
            get;
            set;
        }
    }

    /// <summary>Complete <c>dimension</c> element.</summary>
    internal sealed class XmlDimensionPackage : XmlPackage {
        public string Ref {
            get;
            set;
        }
    }

    /// <summary>
    /// One complete cell fragment (<c>&lt;c&gt;...&lt;/c&gt;</c>) — emit as soon as the element ends.
    /// </summary>
    internal sealed class XmlCellPackage : XmlPackage {
        public string Reference {
            get;
            set;
        }

        public string StyleId {
            get;
            set;
        }

        public string CellType {
            get;
            set;
        }

        public string Formula {
            get;
            set;
        }

        public string ValueText {
            get;
            set;
        }
    }

    /// <summary>
    /// One complete row fragment — emit when <c>&lt;/row&gt;</c> is reached so a renderer
    /// can paint the row without waiting for the rest of the sheet.
    /// </summary>
    internal sealed class XmlRowPackage : XmlPackage {
        public XmlRowPackage() {
            Cells = new List<XmlCellPackage>();
        }

        public string RowIndexAttribute {
            get;
            set;
        }

        public List<XmlCellPackage> Cells {
            get;
            private set;
        }
    }

    /// <summary>Complete <c>mergeCell</c> element.</summary>
    internal sealed class XmlMergePackage : XmlPackage {
        public string Ref {
            get;
            set;
        }
    }

    /// <summary>Complete <c>hyperlink</c> element.</summary>
    internal sealed class XmlHyperlinkPackage : XmlPackage {
        public string Ref {
            get;
            set;
        }

        public string Display {
            get;
            set;
        }

        public string RelationshipId {
            get;
            set;
        }

        public string Location {
            get;
            set;
        }
    }

    /// <summary>Marker that <c>sheetData</c> has ended.</summary>
    internal sealed class XmlSheetDataEndPackage : XmlPackage {
    }
}
