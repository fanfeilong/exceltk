using System;
using System.Collections.Generic;
using System.IO;
using Exceltk.Reader.Binary;
using Exceltk.Reader.Package;

namespace Exceltk.Reader.Parser {
    /// <summary>
    /// Streams <see cref="BinaryPackage"/> entities (one BIFF record each) from a workbook stream.
    /// </summary>
    internal sealed class BinaryPackageParser : XlsStream, IPackageParser<BinaryPackage> {
        private readonly byte[] m_bytes;
        private readonly int m_size;
        private readonly ExcelBinaryReader m_reader;
        private int m_offset;

        public BinaryPackageParser(XlsHeader hdr, uint streamStart, bool isMini, XlsRootDirectory rootDir,
            ExcelBinaryReader reader)
            : base(hdr, streamStart, isMini, rootDir) {
            m_reader = reader;
            m_bytes = base.ReadStream();
            m_size = m_bytes.Length;
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
        /// Read the next BIFF record as a <see cref="BinaryPackage"/> and advance the cursor.
        /// </summary>
        public BinaryPackage Read() {
            if ((uint)m_offset >= m_bytes.Length) {
                return null;
            }

            XlsBiffRecord rec = XlsBiffRecord.GetRecord(m_bytes, (uint)m_offset, m_reader);
            m_offset += rec.Size;

            if (m_offset > m_size) {
                return null;
            }

            return new BinaryPackage(rec);
        }

        /// <summary>
        /// Read a record package at an absolute offset without moving the cursor.
        /// </summary>
        public BinaryPackage ReadAt(int offset) {
            if ((uint)offset >= m_bytes.Length) {
                return null;
            }

            XlsBiffRecord rec = XlsBiffRecord.GetRecord(m_bytes, (uint)offset, m_reader);

            if (m_reader.ReadOption == ReadOption.Strict) {
                if (offset + rec.Size > m_size) {
                    return null;
                }
            }

            return new BinaryPackage(rec);
        }

        /// <summary>
        /// Stream all remaining packages from the current cursor to the end.
        /// </summary>
        public IEnumerable<BinaryPackage> Parse() {
            BinaryPackage package;
            while ((package = Read()) != null) {
                yield return package;
            }
        }
    }
}
