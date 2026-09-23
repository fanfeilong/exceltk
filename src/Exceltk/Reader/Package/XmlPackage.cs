using System.Collections.Generic;

namespace Exceltk.Reader.Package {
    /// <summary>
    /// Base type for OpenXML worksheet semantic units streamed by <c>XmlPackageParser</c>.
    /// </summary>
    internal abstract class XmlPackage : IPackage {
    }

    internal sealed class XmlWorksheetStartPackage : XmlPackage {
        public string NamespaceUri {
            get;
            set;
        }
    }

    internal sealed class XmlDimensionPackage : XmlPackage {
        public string Ref {
            get;
            set;
        }
    }

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

    internal sealed class XmlRowPackage : XmlPackage {
        public XmlRowPackage() {
            Cells = new List<XmlCellPackage>();
        }

        /// <summary>1-based row index from the <c>r</c> attribute, when present.</summary>
        public string RowIndexAttribute {
            get;
            set;
        }

        public List<XmlCellPackage> Cells {
            get;
            private set;
        }
    }

    internal sealed class XmlMergePackage : XmlPackage {
        public string Ref {
            get;
            set;
        }
    }

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

    internal sealed class XmlSheetDataEndPackage : XmlPackage {
    }
}
