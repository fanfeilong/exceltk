using System;
using System.IO;

namespace Exceltk.Reader.Binary {
    /// <summary>
    /// Represents a BIFF stream
    /// </summary>
    internal class XlsBiffStream : XlsStream {
        private readonly byte[] bytes;
        private readonly int m_size;
        private readonly ExcelBinaryReader reader;
        private int m_offset;

        public XlsBiffStream(XlsHeader hdr, uint streamStart, bool isMini, XlsRootDirectory rootDir,
                             ExcelBinaryReader reader)
            : base(hdr, streamStart, isMini, rootDir) {
            this.reader=reader;
            bytes=base.ReadStream();
            m_size=bytes.Length;
            m_offset=0;
        }

        /// <summary>
        /// Returns size of BIFF stream in bytes
        /// </summary>
        public int Size {
            get {
                return m_size;
            }
        }

        /// <summary>
        /// Returns current position in BIFF stream
        /// </summary>
        public int Position {
            get {
                return m_offset;
            }
        }

        /// <summary>
        /// Sets stream pointer to the specified offset
        /// </summary>
        /// <param name="offset">Offset value</param>
        /// <param name="origin">Offset origin</param>
        public void Seek(int offset, SeekOrigin origin) {
            switch (origin) {
                case SeekOrigin.Begin:
                    m_offset=offset;
                    break;
                case SeekOrigin.Current:
                    m_offset+=offset;
                    break;
                case SeekOrigin.End:
                    m_offset=m_size-offset;
                    break;
            }
            
            if (m_offset<0){
                string message = string.Format("{0} On offset={1}", Errors.ErrorBIFFIlegalBefore, offset);
                throw new ArgumentOutOfRangeException(message);
            }

            if (m_offset>m_size){
                string message = string.Format("{0} On offset={1}", Errors.ErrorBIFFIlegalAfter,offset);
                throw new ArgumentOutOfRangeException(message);
            }
        }

        /// <summary>
        /// Reads record under cursor and advances cursor position to next record
        /// </summary>
        /// <returns></returns>
        public XlsBiffRecord Read() {
            if ((uint)m_offset>=bytes.Length){
                return null;
            }

            XlsBiffRecord rec=XlsBiffRecord.GetRecord(bytes, (uint)m_offset, reader);
            m_offset+=rec.Size;
            
            if (m_offset>m_size){
                return null;
            }

            return rec;
        }

        /// <summary>
        /// Reads record at specified offset, does not change cursor position
        /// </summary>
        /// <param name="offset"></param>
        /// <returns></returns>
        public XlsBiffRecord ReadAt(int offset) {
            if ((uint)offset>=bytes.Length){
                return null;
            }

            XlsBiffRecord rec=XlsBiffRecord.GetRecord(bytes, (uint)offset, reader);

            //choose ReadOption.Loose to skip this check (e.g. sql reporting services)
            if (reader.ReadOption==ReadOption.Strict) {
                if (offset+rec.Size>m_size){
                    return null;
                }
            }

            return rec;
        }
    }
}