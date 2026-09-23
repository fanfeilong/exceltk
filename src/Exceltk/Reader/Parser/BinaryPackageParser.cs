using System;
using System.Collections.Generic;
using System.IO;
using Exceltk.Reader.Binary;
using Exceltk.Reader.Package;

namespace Exceltk.Reader.Parser {
    /// <summary>
    /// Streams BIFF <see cref="BinaryPackage"/> entities (each <see cref="XlsBiffRecord"/>)
    /// from a workbook stream. Emits a package as soon as one record's header+body is complete;
    /// each package owns an isolated byte slice so a renderer need not retain the full workbook.
    /// </summary>
    internal sealed class BinaryPackageParser : XlsStream, IPullPackageParser<BinaryPackage> {
        private readonly byte[] m_workbookBytes;
        private readonly int m_size;
        private readonly ExcelBinaryReader m_reader;
        private int m_offset;

        public BinaryPackageParser(XlsHeader hdr, uint streamStart, bool isMini, XlsRootDirectory rootDir,
            ExcelBinaryReader reader)
            : base(hdr, streamStart, isMini, rootDir) {
            m_reader = reader;
            // OLE compound storage still needs sector assembly to reach the workbook stream;
            // once positioned, we only slice one record at a time into each package.
            m_workbookBytes = base.ReadStream();
            m_size = m_workbookBytes.Length;
            m_offset = 0;
        }

        public int Size {
            get {
                return m_size;
            }
        }

        public int Position {
            get {
                return m_offset;
            }
        }

        public void Seek(int offset, SeekOrigin origin) {
            switch (origin) {
                case SeekOrigin.Begin:
                    m_offset = offset;
                    break;
                case SeekOrigin.Current:
                    m_offset += offset;
                    break;
                case SeekOrigin.End:
                    m_offset = m_size - offset;
                    break;
            }

            if (m_offset < 0) {
                string message = string.Format("{0} On offset={1}", Errors.ErrorBIFFIlegalBefore, offset);
                throw new ArgumentOutOfRangeException(message);
            }

            if (m_offset > m_size) {
                string message = string.Format("{0} On offset={1}", Errors.ErrorBIFFIlegalAfter, offset);
                throw new ArgumentOutOfRangeException(message);
            }
        }

        /// <summary>
        /// Read the next complete BIFF record package (local format unit) and advance.
        /// </summary>
        public XlsBiffRecord Read() {
            BinaryPackage package;
            if (!TryRead(out package)) {
                return null;
            }
            return (XlsBiffRecord)package;
        }

        /// <summary>
        /// Peek/build a record package at an absolute workbook-stream offset (no cursor move).
        /// Still allocates a one-record slice — suitable for index/DBCELL jumps.
        /// </summary>
        public XlsBiffRecord ReadAt(int offset) {
            if ((uint)offset >= m_workbookBytes.Length || offset + 4 > m_size) {
                return null;
            }

            ushort recordSize = BitConverter.ToUInt16(m_workbookBytes, offset + 2);
            int size = 4 + recordSize;
            if (m_reader.ReadOption == ReadOption.Strict && offset + size > m_size) {
                return null;
            }
            if (offset + size > m_size) {
                size = m_size - offset;
            }

            byte[] slice = SliceRecord(offset, size);
            return XlsBiffRecord.GetRecord(slice, 0, m_reader);
        }

        public bool TryRead(out BinaryPackage package) {
            package = null;
            if ((uint)m_offset >= m_workbookBytes.Length || m_offset + 4 > m_size) {
                return false;
            }

            ushort recordSize = BitConverter.ToUInt16(m_workbookBytes, m_offset + 2);
            int size = 4 + recordSize;
            if (m_offset + size > m_size) {
                return false;
            }

            byte[] slice = SliceRecord(m_offset, size);
            m_offset += size;
            package = XlsBiffRecord.GetRecord(slice, 0, m_reader);
            return package != null;
        }

        public IEnumerable<BinaryPackage> Parse() {
            BinaryPackage package;
            while (TryRead(out package)) {
                yield return package;
            }
        }

        private byte[] SliceRecord(int offset, int size) {
            var slice = new byte[size];
            Buffer.BlockCopy(m_workbookBytes, offset, slice, 0, size);
            return slice;
        }
    }
}
