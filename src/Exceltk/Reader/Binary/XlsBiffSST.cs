using System;
using System.Collections.Generic;
using System.Text;

namespace Exceltk.Reader.Binary {
    /// <summary>
    /// BIFF Package: Shared String Table (SST).
    /// Members: count, uniqueCount, stringData (body after 8-byte count header).
    /// CONTINUE payloads remain Reader-side (Append + ReadStrings).
    /// </summary>
    internal class XlsBiffSST : BinaryPackage {
        private readonly List<uint> continues = new List<uint>();
        private readonly List<string> m_strings;
        private uint m_size;
        private uint m_count;
        private uint m_uniqueCount;
        private byte[] m_stringData = Array.Empty<byte>();

        internal XlsBiffSST(ExcelBinaryReader reader)
            : base(reader) {
            m_strings = new List<string>();
        }

        internal XlsBiffSST(byte[] bytes, uint offset, ExcelBinaryReader reader)
            : base(bytes, offset, reader) {
            m_size = RecordSize;
            m_strings = new List<string>();
        }

        protected override void DecodeBody(byte[] buffer, int offset, int length) {
            m_size = (uint)Math.Max(0, length);
            if (length < 8) {
                m_count = 0;
                m_uniqueCount = 0;
                m_stringData = Array.Empty<byte>();
                return;
            }

            m_count = BodyReadUInt32(buffer, offset, 0x0);
            m_uniqueCount = BodyReadUInt32(buffer, offset, 0x4);
            int dataLen = length - 8;
            m_stringData = new byte[dataLen];
            if (dataLen > 0) {
                Buffer.BlockCopy(buffer, offset + 8, m_stringData, 0, dataLen);
            }
        }

        protected override int GetRequiredEncodeBodyBufferLength() {
            return 8 + m_stringData.Length;
        }

        protected override void EncodeBody(byte[] buffer, int offset, int capacity, out int written) {
            int need = GetRequiredEncodeBodyBufferLength();
            if (capacity < need) {
                throw new ArgumentException(Errors.ErrorBIFFBufferSize);
            }
            WriteUInt32(buffer, offset + 0x0, m_count);
            WriteUInt32(buffer, offset + 0x4, m_uniqueCount);
            if (m_stringData.Length > 0) {
                Buffer.BlockCopy(m_stringData, 0, buffer, offset + 8, m_stringData.Length);
            }
            written = need;
        }

        public uint Count {
            get {
                return m_count;
            }
        }

        public uint UniqueCount {
            get {
                return m_uniqueCount;
            }
        }

        /// <summary>
        /// Reads strings from BIFF stream into SST array (may span CONTINUE packages).
        /// </summary>
        public void ReadStrings() {
            uint offset = (uint)m_readoffset + 8;
            uint last = (uint)m_readoffset + RecordSize;
            int lastcontinue = 0;
            uint count = UniqueCount;
            while (offset < last) {
                var str = new XlsFormattedUnicodeString(m_bytes, offset);
                uint prefix = str.HeadSize;
                uint postfix = str.TailSize;
                uint len = str.CharacterCount;
                uint size = prefix + postfix + len + ((str.IsMultiByte) ? len : 0);
                if (offset + size > last) {
                    if (lastcontinue >= continues.Count) {
                        break;
                    }

                    uint contoffset = continues[lastcontinue];
                    byte encoding = Buffer.GetByte(m_bytes, (int)contoffset + 4);
                    var buff = new byte[size * 2];
                    Buffer.BlockCopy(m_bytes, (int)offset, buff, 0, (int)(last - offset));
                    if (encoding == 0 && str.IsMultiByte) {
                        len -= (last - prefix - offset) / 2;

                        string temp = Exceltk.Extension.DefaultEncoding().GetString(
                            m_bytes,
                            (int)contoffset + 5,
                            (int)len);

                        byte[] tempbytes = Encoding.Unicode.GetBytes(temp);

                        Buffer.BlockCopy(
                            tempbytes,
                            0,
                            buff,
                            (int)(last - offset),
                            tempbytes.Length);

                        Buffer.BlockCopy(
                            m_bytes,
                            (int)(contoffset + 5 + len),
                            buff,
                            (int)(last - offset + len + len),
                            (int)postfix);

                        offset = contoffset + 5 + len + postfix;

                    } else if (encoding == 1 && str.IsMultiByte == false) {
                        len -= (last - offset - prefix);

                        string temp = Encoding.Unicode.GetString(
                            m_bytes,
                            (int)contoffset + 5,
                            (int)(len + len));

                        byte[] tempbytes = Exceltk.Extension.DefaultEncoding().GetBytes(temp);

                        Buffer.BlockCopy(
                            tempbytes,
                            0,
                            buff,
                            (int)(last - offset),
                            tempbytes.Length);

                        Buffer.BlockCopy(
                            m_bytes,
                            (int)(contoffset + 5 + len + len),
                            buff,
                            (int)(last - offset + len),
                            (int)postfix);

                        offset = contoffset + 5 + len + len + postfix;

                    } else {
                        Buffer.BlockCopy(
                            m_bytes,
                            (int)contoffset + 5,
                            buff,
                            (int)(last - offset),
                            (int)(size - last + offset));

                        offset = contoffset + 5 + size - last + offset;
                    }

                    last = contoffset + 4 + BitConverter.ToUInt16(m_bytes, (int)contoffset + 2);
                    lastcontinue++;

                    str = new XlsFormattedUnicodeString(buff, 0);
                } else {
                    offset += size;
                    if (offset == last) {
                        if (lastcontinue < continues.Count) {
                            uint contoffset = continues[lastcontinue];
                            offset = contoffset + 4;
                            last = offset + BitConverter.ToUInt16(m_bytes, (int)contoffset + 2);
                            lastcontinue++;
                        } else {
                            count = 1;
                        }
                    }
                }
                m_strings.Add(str.Value);
                count--;
                if (count == 0) {
                    break;
                }
            }
        }

        public string GetString(uint SSTIndex) {
            if (SSTIndex < m_strings.Count) {
                return m_strings[(int)SSTIndex];
            }
            return string.Empty;
        }

        public void Append(XlsBiffContinue fragment) {
            continues.Add((uint)fragment.Offset);
            m_size += (uint)fragment.Size;
        }
    }
}
