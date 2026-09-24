using System;
using Exceltk.Reader.Package;

namespace Exceltk.Reader.Binary {
    /// <summary>
    /// BIFF package base (XUdt-style): frames type+size header; subclasses own EncodeBody/DecodeBody.
    /// Each concrete <c>XlsBiff*</c> is one Office BIFF command type.
    /// Unknown types use this base and store a raw body.
    /// </summary>
    internal class BinaryPackage : IPackage {
        protected readonly ExcelBinaryReader reader;
        protected byte[] m_bytes;
        protected int m_readoffset;
        private int m_streamOffset = -1;

        protected BIFFRECORDTYPE m_id;
        protected ushort m_bodyLength;
        /// <summary>Raw body for unknown record types (encode/decode passthrough).</summary>
        protected byte[] m_rawBody;

        public const int HeaderSize = 4;

        public static ushort ReadHeaderType(byte[] header, int offset) {
            return BitConverter.ToUInt16(header, offset);
        }

        public static ushort ReadHeaderBodyLength(byte[] header, int offset) {
            return BitConverter.ToUInt16(header, offset + 2);
        }

        protected BinaryPackage(ExcelBinaryReader reader) {
            this.reader = reader;
            m_bytes = Array.Empty<byte>();
            m_readoffset = 4;
        }

        protected BinaryPackage(BIFFRECORDTYPE id, ExcelBinaryReader reader)
            : this(reader) {
            m_id = id;
        }

        protected BinaryPackage(byte[] bytes, uint offset, ExcelBinaryReader reader)
            : this(reader) {
            if (bytes.Length - offset < 4) {
                throw new ArgumentException(Errors.ErrorBIFFRecordSize);
            }
            AttachSharedBytes(bytes, (int)offset);
        }

        /// <summary>
        /// IPackage entry: decode a complete BIFF record (4-byte header + body).
        /// </summary>
        void IPackage.DecodeBody(byte[] data, int offset, int count) {
            DecodePackage(data, offset, count);
        }

        /// <summary>Decode a complete BIFF record (header + body), then <see cref="DecodeBody"/> for fields.</summary>
        public void DecodePackage(byte[] data, int offset, int count) {
            if (data == null) {
                throw new ArgumentNullException("data");
            }
            if (offset < 0 || count < 0 || offset + count > data.Length) {
                throw new ArgumentOutOfRangeException("count");
            }
            if (count < HeaderSize) {
                throw new ArgumentException(Errors.ErrorBIFFRecordSize);
            }

            m_bytes = new byte[count];
            Buffer.BlockCopy(data, offset, m_bytes, 0, count);
            m_readoffset = HeaderSize;
            m_streamOffset = -1;
            m_id = (BIFFRECORDTYPE)BitConverter.ToUInt16(m_bytes, 0);
            m_bodyLength = BitConverter.ToUInt16(m_bytes, 2);

            if (reader.ReadOption == ReadOption.Strict) {
                if (count < Size) {
                    throw new ArgumentException(Errors.ErrorBIFFBufferSize);
                }
            }

            int bodyLen = Math.Min(m_bodyLength, Math.Max(0, count - HeaderSize));
            DecodeBody(m_bytes, m_readoffset, bodyLen);
        }

        /// <summary>
        /// Attach a shared workbook buffer at an absolute BIFF-stream offset (no copy).
        /// Required for SST + CONTINUE string assembly.
        /// </summary>
        internal void AttachSharedBytes(byte[] bytes, int offset) {
            if (bytes == null) {
                throw new ArgumentNullException("bytes");
            }
            if (offset < 0 || offset + HeaderSize > bytes.Length) {
                throw new ArgumentOutOfRangeException("offset");
            }

            m_bytes = bytes;
            m_readoffset = HeaderSize + offset;
            m_streamOffset = offset;
            m_id = (BIFFRECORDTYPE)BitConverter.ToUInt16(m_bytes, offset);
            m_bodyLength = BitConverter.ToUInt16(m_bytes, offset + 2);

            if (reader.ReadOption == ReadOption.Strict) {
                if (bytes.Length < offset + Size) {
                    throw new ArgumentException(Errors.ErrorBIFFBufferSize);
                }
            }

            int available = bytes.Length - m_readoffset;
            int bodyLen = Math.Min(m_bodyLength, Math.Max(0, available));
            DecodeBody(m_bytes, m_readoffset, bodyLen);
        }

        /// <summary>Body-relative UInt16 from an explicit body buffer (DecodeBody).</summary>
        protected static ushort BodyReadUInt16(byte[] buffer, int bodyOffset, int fieldOfs) {
            return BitConverter.ToUInt16(buffer, bodyOffset + fieldOfs);
        }

        protected static uint BodyReadUInt32(byte[] buffer, int bodyOffset, int fieldOfs) {
            return BitConverter.ToUInt32(buffer, bodyOffset + fieldOfs);
        }

        protected static int BodyReadInt32(byte[] buffer, int bodyOffset, int fieldOfs) {
            return BitConverter.ToInt32(buffer, bodyOffset + fieldOfs);
        }

        protected static long BodyReadInt64(byte[] buffer, int bodyOffset, int fieldOfs) {
            return BitConverter.ToInt64(buffer, bodyOffset + fieldOfs);
        }

        protected static byte BodyReadByte(byte[] buffer, int bodyOffset, int fieldOfs) {
            return buffer[bodyOffset + fieldOfs];
        }

        protected static double BodyReadDouble(byte[] buffer, int bodyOffset, int fieldOfs) {
            return BitConverter.ToDouble(buffer, bodyOffset + fieldOfs);
        }

        /// <summary>Copy body bytes for unknown-type base default only.</summary>
        protected void CaptureRawBody(byte[] buffer, int offset, int length) {
            m_rawBody = new byte[length];
            if (length > 0) {
                Buffer.BlockCopy(buffer, offset, m_rawBody, 0, length);
            }
        }

        /// <summary>Write <see cref="m_rawBody"/> — unknown-type base default only.</summary>
        protected void EncodeRawBody(byte[] buffer, int offset, int capacity, out int written) {
            written = 0;
            if (m_rawBody == null || m_rawBody.Length == 0) {
                return;
            }
            if (capacity < m_rawBody.Length) {
                throw new ArgumentException(Errors.ErrorBIFFBufferSize);
            }
            Buffer.BlockCopy(m_rawBody, 0, buffer, offset, m_rawBody.Length);
            written = m_rawBody.Length;
        }

        /// <summary>
        /// Decode Office-defined fields from the BIFF body (excluding the 4-byte header).
        /// Unknown types store opaque <see cref="m_rawBody"/>.
        /// </summary>
        protected virtual void DecodeBody(byte[] buffer, int offset, int length) {
            CaptureRawBody(buffer, offset, length);
        }

        /// <summary>Encode Office-defined body fields (excluding the 4-byte header).</summary>
        protected virtual void EncodeBody(byte[] buffer, int offset, int capacity, out int written) {
            EncodeRawBody(buffer, offset, capacity, out written);
        }

        /// <summary>Required body length for <see cref="EncodeBody"/>.</summary>
        protected virtual int GetRequiredEncodeBodyBufferLength() {
            return m_rawBody != null ? m_rawBody.Length : m_bodyLength;
        }

        /// <summary>
        /// Encode a complete BIFF record: type + size header, then <see cref="EncodeBody"/>.
        /// </summary>
        public bool EncodePackage(byte[] buffer, int offset, int count, out int written) {
            written = 0;
            int bodyLen = GetRequiredEncodeBodyBufferLength();
            if (bodyLen < 0 || bodyLen > ushort.MaxValue) {
                return false;
            }
            int total = HeaderSize + bodyLen;
            if (count < total) {
                return false;
            }

            BitConverter.GetBytes((ushort)m_id).CopyTo(buffer, offset);
            BitConverter.GetBytes((ushort)bodyLen).CopyTo(buffer, offset + 2);
            int bodyWritten;
            EncodeBody(buffer, offset + HeaderSize, bodyLen, out bodyWritten);
            if (bodyWritten != bodyLen) {
                BitConverter.GetBytes((ushort)bodyWritten).CopyTo(buffer, offset + 2);
            }
            written = HeaderSize + bodyWritten;
            m_bodyLength = (ushort)bodyWritten;
            return true;
        }

        internal void AttachStreamOffset(int streamOffset) {
            m_streamOffset = streamOffset;
        }

        internal void SetRecordType(BIFFRECORDTYPE id) {
            m_id = id;
        }

        internal byte[] Bytes {
            get {
                return m_bytes;
            }
        }

        internal int Offset {
            get {
                if (m_streamOffset >= 0) {
                    return m_streamOffset;
                }
                return m_readoffset - HeaderSize;
            }
        }

        public BIFFRECORDTYPE ID {
            get {
                return m_id;
            }
        }

        public ushort RecordSize {
            get {
                return m_bodyLength;
            }
        }

        public int Size {
            get {
                return HeaderSize + m_bodyLength;
            }
        }

        public bool IsCell {
            get {
                switch (ID) {
                    case BIFFRECORDTYPE.FORMULA:
                    case BIFFRECORDTYPE.BLANK:
                    case BIFFRECORDTYPE.MULBLANK:
                    case BIFFRECORDTYPE.RK:
                    case BIFFRECORDTYPE.MULRK:
                    case BIFFRECORDTYPE.BOOLERR:
                    case BIFFRECORDTYPE.NUMBER:
                    case BIFFRECORDTYPE.LABELSST:
                        return true;
                    default:
                        return false;
                }
            }
        }

        public static BinaryPackage CreateEmpty(BIFFRECORDTYPE id, ExcelBinaryReader reader) {
            BinaryPackage package;
            switch (id) {
                case BIFFRECORDTYPE.BOF_V2:
                case BIFFRECORDTYPE.BOF_V3:
                case BIFFRECORDTYPE.BOF_V4:
                case BIFFRECORDTYPE.BOF:
                    package = new XlsBiffBOF(reader);
                    break;
                case BIFFRECORDTYPE.EOF:
                    package = new XlsBiffEOF(reader);
                    break;
                case BIFFRECORDTYPE.INTERFACEHDR:
                    package = new XlsBiffInterfaceHdr(reader);
                    break;
                case BIFFRECORDTYPE.SST:
                    package = new XlsBiffSST(reader);
                    break;
                case BIFFRECORDTYPE.INDEX:
                    package = new XlsBiffIndex(reader);
                    break;
                case BIFFRECORDTYPE.ROW:
                    package = new XlsBiffRow(reader);
                    break;
                case BIFFRECORDTYPE.DBCELL:
                    package = new XlsBiffDbCell(reader);
                    break;
                case BIFFRECORDTYPE.BOOLERR:
                case BIFFRECORDTYPE.BOOLERR_OLD:
                case BIFFRECORDTYPE.BLANK:
                case BIFFRECORDTYPE.BLANK_OLD:
                    package = new XlsBiffBlankCell(reader);
                    break;
                case BIFFRECORDTYPE.MULBLANK:
                    package = new XlsBiffMulBlankCell(reader);
                    break;
                case BIFFRECORDTYPE.LABEL:
                case BIFFRECORDTYPE.LABEL_OLD:
                case BIFFRECORDTYPE.RSTRING:
                    package = new XlsBiffLabelCell(reader);
                    break;
                case BIFFRECORDTYPE.LABELSST:
                    package = new XlsBiffLabelSSTCell(reader);
                    break;
                case BIFFRECORDTYPE.INTEGER:
                case BIFFRECORDTYPE.INTEGER_OLD:
                    package = new XlsBiffIntegerCell(reader);
                    break;
                case BIFFRECORDTYPE.NUMBER:
                case BIFFRECORDTYPE.NUMBER_OLD:
                    package = new XlsBiffNumberCell(reader);
                    break;
                case BIFFRECORDTYPE.RK:
                    package = new XlsBiffRKCell(reader);
                    break;
                case BIFFRECORDTYPE.MULRK:
                    package = new XlsBiffMulRKCell(reader);
                    break;
                case BIFFRECORDTYPE.FORMULA:
                case BIFFRECORDTYPE.FORMULA_OLD:
                    package = new XlsBiffFormulaCell(reader);
                    break;
                case BIFFRECORDTYPE.FORMAT_V23:
                case BIFFRECORDTYPE.FORMAT:
                    package = new XlsBiffFormatString(reader);
                    break;
                case BIFFRECORDTYPE.STRING:
                    package = new XlsBiffFormulaString(reader);
                    break;
                case BIFFRECORDTYPE.CONTINUE:
                    package = new XlsBiffContinue(reader);
                    break;
                case BIFFRECORDTYPE.DIMENSIONS:
                    package = new XlsBiffDimensions(reader);
                    break;
                case BIFFRECORDTYPE.BOUNDSHEET:
                    package = new XlsBiffBoundSheet(reader);
                    break;
                case BIFFRECORDTYPE.WINDOW1:
                    package = new XlsBiffWindow1(reader);
                    break;
                case BIFFRECORDTYPE.CODEPAGE:
                case BIFFRECORDTYPE.FNGROUPCOUNT:
                case BIFFRECORDTYPE.RECORD1904:
                case BIFFRECORDTYPE.BOOKBOOL:
                case BIFFRECORDTYPE.BACKUP:
                case BIFFRECORDTYPE.HIDEOBJ:
                case BIFFRECORDTYPE.USESELFS:
                    package = new XlsBiffSimpleValueRecord(reader);
                    break;
                case BIFFRECORDTYPE.UNCALCED:
                    package = new XlsBiffUncalced(reader);
                    break;
                case BIFFRECORDTYPE.QUICKTIP:
                    package = new XlsBiffQuickTip(reader);
                    break;
                case BIFFRECORDTYPE.HLINK:
                    package = new XlsBiffHyperLink(reader);
                    break;
                default:
                    package = new BinaryPackage(reader);
                    break;
            }
            package.m_id = id;
            return package;
        }

        public static BinaryPackage GetRecord(byte[] bytes, uint offset, ExcelBinaryReader reader) {
            if (offset >= bytes.Length) {
                return null;
            }
            if (bytes.Length - offset < HeaderSize) {
                return null;
            }

            ushort id = BitConverter.ToUInt16(bytes, (int)offset);
            ushort recordSize = BitConverter.ToUInt16(bytes, (int)offset + 2);
            int size = HeaderSize + recordSize;

            if (reader.ReadOption == ReadOption.Strict) {
                if (offset + size > bytes.Length) {
                    return null;
                }
            }

            BinaryPackage record = CreateEmpty((BIFFRECORDTYPE)id, reader);
            record.AttachSharedBytes(bytes, (int)offset);
            return record;
        }

        /// <summary>Read a body-relative byte (requires m_readoffset at body start).</summary>
        public byte ReadByte(int offset) {
            return Buffer.GetByte(m_bytes, m_readoffset + offset);
        }

        public ushort ReadUInt16(int offset) {
            return BitConverter.ToUInt16(m_bytes, m_readoffset + offset);
        }

        public uint ReadUInt32(int offset) {
            return BitConverter.ToUInt32(m_bytes, m_readoffset + offset);
        }

        public ulong ReadUInt64(int offset) {
            return BitConverter.ToUInt64(m_bytes, m_readoffset + offset);
        }

        public short ReadInt16(int offset) {
            return BitConverter.ToInt16(m_bytes, m_readoffset + offset);
        }

        public int ReadInt32(int offset) {
            return BitConverter.ToInt32(m_bytes, m_readoffset + offset);
        }

        public long ReadInt64(int offset) {
            return BitConverter.ToInt64(m_bytes, m_readoffset + offset);
        }

        public byte[] ReadArray(int offset, int size) {
            var tmp = new byte[size];
            Buffer.BlockCopy(m_bytes, m_readoffset + offset, tmp, 0, size);
            return tmp;
        }

        public float ReadFloat(int offset) {
            return BitConverter.ToSingle(m_bytes, m_readoffset + offset);
        }

        public double ReadDouble(int offset) {
            return BitConverter.ToDouble(m_bytes, m_readoffset + offset);
        }

        protected static void WriteUInt16(byte[] buffer, int offset, ushort value) {
            buffer[offset] = (byte)(value & 0xFF);
            buffer[offset + 1] = (byte)((value >> 8) & 0xFF);
        }

        protected static void WriteUInt32(byte[] buffer, int offset, uint value) {
            WriteUInt16(buffer, offset, (ushort)(value & 0xFFFF));
            WriteUInt16(buffer, offset + 2, (ushort)((value >> 16) & 0xFFFF));
        }

        protected static void WriteInt32(byte[] buffer, int offset, int value) {
            WriteUInt32(buffer, offset, (uint)value);
        }

        protected static void WriteInt64(byte[] buffer, int offset, long value) {
            WriteUInt32(buffer, offset, (uint)(value & 0xFFFFFFFF));
            WriteUInt32(buffer, offset + 4, (uint)((ulong)value >> 32));
        }

        protected static void WriteDouble(byte[] buffer, int offset, double value) {
            byte[] bits = BitConverter.GetBytes(value);
            Buffer.BlockCopy(bits, 0, buffer, offset, 8);
        }
    }
}
