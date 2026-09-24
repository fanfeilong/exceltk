namespace Exceltk.Reader.Binary {
    /// <summary>
    /// BIFF Package: BOF (beginning of file/workbook/worksheet stream).
    /// </summary>
    internal class XlsBiffBOF : BinaryPackage {
        private ushort m_version;
        private BIFFTYPE m_type;
        private ushort m_creationId;
        private ushort m_creationYear;
        private uint m_historyFlag;
        private uint m_minVersionToOpen;

        internal XlsBiffBOF(ExcelBinaryReader reader)
            : base(reader) {
        }

        internal XlsBiffBOF(byte[] bytes, uint offset, ExcelBinaryReader reader)
            : base(bytes, offset, reader) {
        }

        protected override void DecodeBody(byte[] buffer, int offset, int length) {
            m_version = ReadUInt16(0x0);
            m_type = (BIFFTYPE)ReadUInt16(0x2);
            m_creationId = RecordSize < 6 ? (ushort)0 : ReadUInt16(0x4);
            m_creationYear = RecordSize < 8 ? (ushort)0 : ReadUInt16(0x6);
            m_historyFlag = RecordSize < 12 ? 0u : ReadUInt32(0x8);
            m_minVersionToOpen = RecordSize < 16 ? 0u : ReadUInt32(0xC);
        }

        protected override int GetRequiredEncodeBodyBufferLength() {
            if (m_minVersionToOpen != 0 || m_historyFlag != 0) return 16;
            if (m_creationYear != 0 || m_creationId != 0) return 8;
            return 4;
        }

        protected override void EncodeBody(byte[] buffer, int offset, int capacity, out int written) {
            int need = GetRequiredEncodeBodyBufferLength();
            if (capacity < need) throw new System.ArgumentException(Errors.ErrorBIFFBufferSize);
            WriteUInt16(buffer, offset + 0x0, m_version);
            WriteUInt16(buffer, offset + 0x2, (ushort)m_type);
            if (need >= 6) WriteUInt16(buffer, offset + 0x4, m_creationId);
            if (need >= 8) WriteUInt16(buffer, offset + 0x6, m_creationYear);
            if (need >= 12) WriteUInt32(buffer, offset + 0x8, m_historyFlag);
            if (need >= 16) WriteUInt32(buffer, offset + 0xC, m_minVersionToOpen);
            written = need;
        }

        public ushort Version {
            get {
                return m_version;
            }
        }

        public BIFFTYPE Type {
            get {
                return m_type;
            }
        }

        /// <remarks>Not used before BIFF5</remarks>
        public ushort CreationID {
            get {
                return m_creationId;
            }
        }

        /// <remarks>Not used before BIFF5</remarks>
        public ushort CreationYear {
            get {
                return m_creationYear;
            }
        }

        /// <remarks>Not used before BIFF8</remarks>
        public uint HistoryFlag {
            get {
                return m_historyFlag;
            }
        }

        /// <remarks>Not used before BIFF8</remarks>
        public uint MinVersionToOpen {
            get {
                return m_minVersionToOpen;
            }
        }
    }
}
