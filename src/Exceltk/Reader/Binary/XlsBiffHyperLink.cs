using System;
using System.Text;

namespace Exceltk.Reader.Binary {
    /// <summary>
    /// BIFF Package: HLINK (hyperlink).
    /// Field-level: cell range, GUID, flags, reserved dword, optional description/frame, URL.
    /// </summary>
    internal class XlsBiffHyperLink : BinaryPackage {
        private short m_firstRow;
        private short m_lastRow;
        private short m_firstCol;
        private short m_lastCol;
        private byte[] m_guid = new byte[16];
        private uint m_flags;
        /// <summary>Body+28..31 (between flags and optional/URL blocks).</summary>
        private uint m_reservedAfterFlags;
        private byte[] m_descriptionBlock = Array.Empty<byte>();
        private byte[] m_frameBlock = Array.Empty<byte>();
        private string m_url = string.Empty;
        private XlsCellRangeAddress m_cellRangeAddress;

        internal XlsBiffHyperLink(ExcelBinaryReader reader)
            : base(reader) {
        }

        internal XlsBiffHyperLink(byte[] bytes, uint offset, ExcelBinaryReader reader)
            : base(bytes, offset, reader) {
        }

        protected override void DecodeBody(byte[] buffer, int offset, int length) {
            if (length < 32) {
                m_flags = 0;
                m_reservedAfterFlags = 0;
                m_url = string.Empty;
                m_descriptionBlock = Array.Empty<byte>();
                m_frameBlock = Array.Empty<byte>();
                m_cellRangeAddress = new XlsCellRangeAddress(buffer, offset);
                return;
            }

            m_firstRow = (short)BodyReadUInt16(buffer, offset, 0);
            m_lastRow = (short)BodyReadUInt16(buffer, offset, 2);
            m_firstCol = (short)BodyReadUInt16(buffer, offset, 4);
            m_lastCol = (short)BodyReadUInt16(buffer, offset, 6);
            Buffer.BlockCopy(buffer, offset + 8, m_guid, 0, 16);
            // Flags at body+24 (== absolute Offset+28 when Offset is record start).
            m_flags = BodyReadUInt32(buffer, offset, 24);
            m_reservedAfterFlags = BodyReadUInt32(buffer, offset, 28);
            m_cellRangeAddress = new XlsCellRangeAddress(buffer, offset);

            // Optional/URL blocks start at body+32 (legacy: 32 + m_readoffset).
            int pos = 32;
            m_descriptionBlock = Array.Empty<byte>();
            m_frameBlock = Array.Empty<byte>();

            if (HasDescription && pos + 4 <= length) {
                int descriptCount = (int)BodyReadUInt32(buffer, offset, pos);
                int blockLen = 4 + Math.Max(0, descriptCount);
                blockLen = Math.Min(blockLen, length - pos);
                m_descriptionBlock = new byte[blockLen];
                Buffer.BlockCopy(buffer, offset + pos, m_descriptionBlock, 0, blockLen);
                pos += blockLen;
            }

            if (HasTatgetFrame && pos + 4 <= length) {
                int frameCount = (int)BodyReadUInt32(buffer, offset, pos);
                int blockLen = 4 + Math.Max(0, frameCount);
                blockLen = Math.Min(blockLen, length - pos);
                m_frameBlock = new byte[blockLen];
                Buffer.BlockCopy(buffer, offset + pos, m_frameBlock, 0, blockLen);
                pos += blockLen;
            }

            m_url = string.Empty;
            if (pos + 4 <= length) {
                int urlSize = (int)BodyReadUInt32(buffer, offset, pos);
                pos += 4;
                int charBytes = 2 * Math.Max(0, urlSize - 1);
                charBytes = Math.Min(charBytes, Math.Max(0, length - pos));
                if (charBytes > 0) {
                    m_url = Encoding.Unicode.GetString(buffer, offset + pos, charBytes);
                }
            }
        }

        protected override int GetRequiredEncodeBodyBufferLength() {
            int urlChars = m_url != null ? m_url.Length : 0;
            int urlBlock = 4 + urlChars * 2;
            return 32 + m_descriptionBlock.Length + m_frameBlock.Length + urlBlock;
        }

        protected override void EncodeBody(byte[] buffer, int offset, int capacity, out int written) {
            int need = GetRequiredEncodeBodyBufferLength();
            if (capacity < need) {
                throw new ArgumentException(Errors.ErrorBIFFBufferSize);
            }

            WriteUInt16(buffer, offset + 0, (ushort)m_firstRow);
            WriteUInt16(buffer, offset + 2, (ushort)m_lastRow);
            WriteUInt16(buffer, offset + 4, (ushort)m_firstCol);
            WriteUInt16(buffer, offset + 6, (ushort)m_lastCol);
            Buffer.BlockCopy(m_guid, 0, buffer, offset + 8, 16);
            WriteUInt32(buffer, offset + 24, m_flags);
            WriteUInt32(buffer, offset + 28, m_reservedAfterFlags);

            int pos = 32;
            if (m_descriptionBlock.Length > 0) {
                Buffer.BlockCopy(m_descriptionBlock, 0, buffer, offset + pos, m_descriptionBlock.Length);
                pos += m_descriptionBlock.Length;
            }
            if (m_frameBlock.Length > 0) {
                Buffer.BlockCopy(m_frameBlock, 0, buffer, offset + pos, m_frameBlock.Length);
                pos += m_frameBlock.Length;
            }

            int urlChars = m_url != null ? m_url.Length : 0;
            WriteUInt32(buffer, offset + pos, (uint)(urlChars + 1));
            pos += 4;
            if (urlChars > 0) {
                byte[] urlBytes = Encoding.Unicode.GetBytes(m_url);
                Buffer.BlockCopy(urlBytes, 0, buffer, offset + pos, urlBytes.Length);
                pos += urlBytes.Length;
            }
            written = pos;
        }

        public UInt32 Flags {
            get {
                return m_flags;
            }
        }

        public bool HasUrl {
            get {
                return (Flags & 0x00000001) == 1;
            }
        }

        public bool IsRelative {
            get {
                return (Flags & 0x00000002) >> 1 == 1;
            }
        }

        public bool HasDescription {
            get {
                uint bit = (Flags & 0x00000014);
                return bit >> 2 == 1 && bit >> 4 == 1;
            }
        }

        public bool HasTextMark {
            get {
                return (Flags & 0x00000008) >> 3 == 1;
            }
        }

        public bool HasTatgetFrame {
            get {
                return (Flags & 0x00000080) >> 7 == 1;
            }
        }

        public bool IsUNC {
            get {
                return (Flags & 0x00000100) >> 8 == 1;
            }
        }

        public bool IsFileOrUrl {
            get {
                return !IsUNC;
            }
        }

        public string Url {
            get {
                return m_url ?? string.Empty;
            }
        }

        public XlsCellRangeAddress CellRangeAddress {
            get {
                return m_cellRangeAddress;
            }
        }
    }
}
