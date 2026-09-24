using System;

namespace Exceltk.Reader.Binary {
    /// <summary>
    /// Represents Workbook's global window description
    /// </summary>
    internal class XlsBiffWindow1 : BinaryPackage {
        #region Window1Flags enum

        [Flags]
        public enum Window1Flags : ushort {
            Hidden=0x1,
            Minimized=0x2,
            //(Reserved) = 0x4,

            HScrollVisible=0x8,
            VScrollVisible=0x10,
            WorkbookTabs=0x20
            //(Other bits are reserved)
        }

        #endregion

        private ushort m_left;
        private ushort m_top;
        private ushort m_width;
        private ushort m_height;
        private Window1Flags m_flags;
        private ushort m_activeTab;
        private ushort m_firstVisibleTab;
        private ushort m_selectedTabCount;
        private ushort m_tabRatio;

        internal XlsBiffWindow1(ExcelBinaryReader reader)
            : base(reader) {
        }

        internal XlsBiffWindow1(byte[] bytes, uint offset, ExcelBinaryReader reader)
            : base(bytes, offset, reader) {
        }

        protected override void DecodeBody(byte[] buffer, int offset, int length) {
            m_left = ReadUInt16(0x0);
            m_top = ReadUInt16(0x2);
            m_width = ReadUInt16(0x4);
            m_height = ReadUInt16(0x6);
            m_flags = (Window1Flags)ReadUInt16(0x8);
            m_activeTab = ReadUInt16(0xA);
            m_firstVisibleTab = ReadUInt16(0xC);
            m_selectedTabCount = ReadUInt16(0xE);
            m_tabRatio = ReadUInt16(0x10);
        }

        protected override int GetRequiredEncodeBodyBufferLength() {
            return 18;
        }

        protected override void EncodeBody(byte[] buffer, int offset, int capacity, out int written) {
            if (capacity < 18) throw new System.ArgumentException(Errors.ErrorBIFFBufferSize);
            WriteUInt16(buffer, offset + 0x0, m_left);
            WriteUInt16(buffer, offset + 0x2, m_top);
            WriteUInt16(buffer, offset + 0x4, m_width);
            WriteUInt16(buffer, offset + 0x6, m_height);
            WriteUInt16(buffer, offset + 0x8, (ushort)m_flags);
            WriteUInt16(buffer, offset + 0xA, m_activeTab);
            WriteUInt16(buffer, offset + 0xC, m_firstVisibleTab);
            WriteUInt16(buffer, offset + 0xE, m_selectedTabCount);
            WriteUInt16(buffer, offset + 0x10, m_tabRatio);
            written = 18;
        }

        /// <summary>
        /// Returns X position of a window
        /// </summary>
        public ushort Left {
            get {
                return m_left;
            }
        }

        /// <summary>
        /// Returns Y position of a window
        /// </summary>
        public ushort Top {
            get {
                return m_top;
            }
        }

        /// <summary>
        /// Returns width of a window
        /// </summary>
        public ushort Width {
            get {
                return m_width;
            }
        }

        /// <summary>
        /// Returns height of a window
        /// </summary>
        public ushort Height {
            get {
                return m_height;
            }
        }

        /// <summary>
        /// Returns window flags
        /// </summary>
        public Window1Flags Flags {
            get {
                return m_flags;
            }
        }

        /// <summary>
        /// Returns active workbook tab (zero-based)
        /// </summary>
        public ushort ActiveTab {
            get {
                return m_activeTab;
            }
        }

        /// <summary>
        /// Returns first visible workbook tab (zero-based)
        /// </summary>
        public ushort FirstVisibleTab {
            get {
                return m_firstVisibleTab;
            }
        }

        /// <summary>
        /// Returns number of selected workbook tabs
        /// </summary>
        public ushort SelectedTabCount {
            get {
                return m_selectedTabCount;
            }
        }

        /// <summary>
        /// Returns workbook tab width to horizontal scrollbar width
        /// </summary>
        public ushort TabRatio {
            get {
                return m_tabRatio;
            }
        }
    }
}
