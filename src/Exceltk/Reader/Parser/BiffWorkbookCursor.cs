using System;
using System.IO;
using Exceltk.Reader.Binary;

namespace Exceltk.Reader.Parser {
    /// <summary>
    /// Random-access BIFF cursor over an assembled workbook stream buffer.
    /// Separate from <see cref="BinaryPackageParser"/> PushData framing (DBCELL / Seek paths).
    /// </summary>
    internal sealed class BiffWorkbookCursor {
        private readonly byte[] m_bytes;
        private readonly ExcelBinaryReader m_reader;
        private int m_offset;

        public BiffWorkbookCursor(byte[] workbookBytes, ExcelBinaryReader reader) {
            m_bytes = workbookBytes ?? throw new ArgumentNullException("workbookBytes");
            m_reader = reader;
            m_offset = 0;
        }

        public byte[] Bytes {
            get {
                return m_bytes;
            }
        }

        public int Size {
            get {
                return m_bytes.Length;
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
                    m_offset = Size - offset;
                    break;
            }

            if (m_offset < 0) {
                string message = string.Format("{0} On offset={1}", Errors.ErrorBIFFIlegalBefore, offset);
                throw new ArgumentOutOfRangeException(message);
            }

            if (m_offset > Size) {
                string message = string.Format("{0} On offset={1}", Errors.ErrorBIFFIlegalAfter, offset);
                throw new ArgumentOutOfRangeException(message);
            }
        }

        /// <summary>Frame the next record at the cursor and advance.</summary>
        public BinaryPackage Read() {
            BinaryPackage record = ReadAt(m_offset);
            if (record == null) {
                return null;
            }
            m_offset += record.Size;
            if (m_offset > Size) {
                m_offset = Size;
            }
            return record;
        }

        /// <summary>Frame one record at an absolute workbook-stream offset (cursor unchanged).</summary>
        public BinaryPackage ReadAt(int offset) {
            return BinaryPackage.GetRecord(m_bytes, (uint)offset, m_reader);
        }
    }
}
