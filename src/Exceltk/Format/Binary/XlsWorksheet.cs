namespace ExcelToolKit.Format.Binary {
    /// <summary>
    /// Represents Worksheet section in workbook
    /// </summary>
    internal class XlsWorksheet {
        private readonly int m_Index;
        private readonly string m_Name=string.Empty;
        private readonly uint m_dataOffset;

        public XlsWorksheet(int index, XlsBiffBoundSheet refSheet) {
            m_Index=index;
            m_Name=refSheet.SheetName;
            m_dataOffset=refSheet.StartOffset;
        }

        /// <summary>
        /// Name of worksheet
        /// </summary>
        public string Name {
            get {
                return m_Name;
            }
        }

        /// <summary>
        /// Zero-based index of worksheet
        /// </summary>
        public int Index {
            get {
                return m_Index;
            }
        }

        /// <summary>
        /// Offset of worksheet data
        /// </summary>
        public uint DataOffset {
            get {
                return m_dataOffset;
            }
        }

        public XlsBiffSimpleValueRecord CalcMode {
            get;
            set;
        }

        public XlsBiffSimpleValueRecord CalcCount {
            get;
            set;
        }

        public XlsBiffSimpleValueRecord RefMode {
            get;
            set;
        }

        public XlsBiffSimpleValueRecord Iteration {
            get;
            set;
        }

        public XlsBiffRecord Delta {
            get;
            set;
        }

        /// <summary>
        /// Dimensions of worksheet
        /// </summary>
        public XlsBiffDimensions Dimensions {
            get;
            set;
        }

        public XlsBiffRecord Window {
            get;
            set;
        }
    }
}