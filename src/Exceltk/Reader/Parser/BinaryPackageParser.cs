using System;
using System.Collections.Generic;
using Exceltk.Reader.Binary;

namespace Exceltk.Reader.Parser {
    /// <summary>
    /// BIFF record framer over a byte stream feed (XUdt-style PushData state machine).
    /// Does not load OLE compound storage — caller assembles the workbook stream and pushes bytes.
    /// </summary>
    internal sealed class BinaryPackageParser : IPackageParser<BinaryPackage> {
        private enum FrameState {
            NeedHeader,
            NeedBody
        }

        private readonly ExcelBinaryReader m_reader;
        private readonly Queue<BinaryPackage> m_queue = new Queue<BinaryPackage>();
        private readonly List<byte> m_pending = new List<byte>(256);
        private readonly byte[] m_header = new byte[BinaryPackage.HeaderSize];

        private FrameState m_state = FrameState.NeedHeader;
        private ushort m_bodyLength;
        private BinaryPackage m_building;
        private int m_streamOffset;
        private int m_recordStartOffset;

        public BinaryPackageParser(ExcelBinaryReader reader) {
            m_reader = reader;
        }

        public bool HavePackage {
            get {
                return m_queue.Count > 0;
            }
        }

        public bool HaveUnParsedData {
            get {
                return m_pending.Count > 0 || m_building != null;
            }
        }

        public void Reset() {
            m_queue.Clear();
            m_pending.Clear();
            m_state = FrameState.NeedHeader;
            m_bodyLength = 0;
            m_building = null;
            m_streamOffset = 0;
            m_recordStartOffset = 0;
        }

        public BinaryPackage PopPackage() {
            return m_queue.Dequeue();
        }

        public void PushData(byte[] data, int offset, int count) {
            if (data == null) {
                throw new ArgumentNullException("data");
            }
            if (offset < 0 || count < 0 || offset + count > data.Length) {
                throw new ArgumentOutOfRangeException("count");
            }

            for (int i = 0; i < count; i++) {
                m_pending.Add(data[offset + i]);
            }

            FrameAvailable();
        }

        private void FrameAvailable() {
            while (true) {
                if (m_state == FrameState.NeedHeader) {
                    if (m_pending.Count < BinaryPackage.HeaderSize) {
                        return;
                    }

                    for (int i = 0; i < BinaryPackage.HeaderSize; i++) {
                        m_header[i] = m_pending[i];
                    }
                    m_pending.RemoveRange(0, BinaryPackage.HeaderSize);

                    var recordType = (BIFFRECORDTYPE)BinaryPackage.ReadHeaderType(m_header, 0);
                    m_bodyLength = BinaryPackage.ReadHeaderBodyLength(m_header, 0);
                    m_recordStartOffset = m_streamOffset;
                    m_streamOffset += BinaryPackage.HeaderSize;
                    m_building = BinaryPackage.CreateEmpty(recordType, m_reader);
                    m_state = FrameState.NeedBody;
                }

                if (m_state == FrameState.NeedBody) {
                    if (m_pending.Count < m_bodyLength) {
                        return;
                    }

                    byte[] body = ConsumePending(m_bodyLength);
                    m_streamOffset += m_bodyLength;

                    var recordBytes = new byte[BinaryPackage.HeaderSize + m_bodyLength];
                    Buffer.BlockCopy(m_header, 0, recordBytes, 0, BinaryPackage.HeaderSize);
                    if (m_bodyLength > 0) {
                        Buffer.BlockCopy(body, 0, recordBytes, BinaryPackage.HeaderSize, m_bodyLength);
                    }

                    m_building.DecodePackage(recordBytes, 0, recordBytes.Length);
                    m_building.AttachStreamOffset(m_recordStartOffset);
                    m_queue.Enqueue(m_building);

                    m_building = null;
                    m_state = FrameState.NeedHeader;
                }
            }
        }

        private byte[] ConsumePending(int count) {
            var result = new byte[count];
            for (int i = 0; i < count; i++) {
                result[i] = m_pending[i];
            }
            m_pending.RemoveRange(0, count);
            return result;
        }
    }
}
