using System;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using Exceltk.Reader;
using Exceltk.Reader.Binary;
using Exceltk.Reader.Package;
using Exceltk.Reader.Parser;
using Exceltk.Reader.Xml;

namespace Exceltk.Test {
    /// <summary>
    /// End-to-end TCP demo: server sends Excel payload in small chunks; client uses
    /// streaming package parsers and prints packages as they complete.
    /// </summary>
    internal static class TcpStreamingDemo {
        private const int ChunkSize = 96;
        private const int ChunkDelayMs = 2;

        public static int Run(string xlsxPath, string xlsPath) {
            Console.WriteLine("=== TCP streaming package demo ===");

            bool xmlOk = RunXmlSheetDemo(xlsxPath);
            bool biffOk = RunBiffPackageDemo(xlsPath);

            if (xmlOk && biffOk) {
                Console.WriteLine("TCP stream test OK");
                Console.WriteLine("Done!");
                return 0;
            }

            Console.WriteLine("TCP stream test FAILED (xml={0}, biff={1})", xmlOk, biffOk);
            return 1;
        }

        /// <summary>
        /// Server streams worksheet XML bytes; client XmlPackageParser emits row/dimension packages live.
        /// </summary>
        private static bool RunXmlSheetDemo(string xlsxPath) {
            Console.WriteLine("-- XML sheet over TCP: {0}", Path.GetFileName(xlsxPath));
            byte[] sheetXml = LoadFirstWorksheetXml(xlsxPath);
            if (sheetXml == null || sheetXml.Length == 0) {
                Console.WriteLine("ERROR: could not load worksheet XML");
                return false;
            }

            var listener = new TcpListener(IPAddress.Loopback, 0);
            listener.Start();
            int port = ((IPEndPoint)listener.LocalEndpoint).Port;

            int rowsSeen = 0;
            int dimSeen = 0;
            Exception clientError = null;

            var clientTask = Task.Run(() => {
                try {
                    using (var client = new TcpClient()) {
                        client.Connect(IPAddress.Loopback, port);
                        using (NetworkStream net = client.GetStream()) {
                            var settings = new XmlReaderSettings {
                                CloseInput = false,
                                IgnoreComments = true,
                                IgnoreWhitespace = false
                            };
                            using (XmlReader xmlReader = XmlReader.Create(net, settings)) {
                                var parser = new XmlPackageParser(xmlReader);
                                foreach (XmlPackage package in parser.Parse()) {
                                    if (package is XmlDimensionPackage dim) {
                                        dimSeen++;
                                        Console.WriteLine("[xml-stream] dimension ref={0}", dim.Ref);
                                    } else if (package is XmlRowPackage row) {
                                        rowsSeen++;
                                        if (rowsSeen <= 15 || rowsSeen % 25 == 0) {
                                            Console.WriteLine("[xml-stream] row#{0} cells={1} r={2}",
                                                rowsSeen, row.Cells.Count, row.RowIndexAttribute);
                                        }
                                    } else if (package is XmlSheetDataEndPackage) {
                                        Console.WriteLine("[xml-stream] sheetData end");
                                    } else if (package is XmlWorksheetStartPackage ws) {
                                        Console.WriteLine("[xml-stream] worksheet ns={0}", ws.NamespaceUri);
                                    }
                                }
                            }
                        }
                    }
                } catch (Exception ex) {
                    clientError = ex;
                }
            });

            using (TcpClient serverClient = listener.AcceptTcpClient())
            using (NetworkStream serverNet = serverClient.GetStream()) {
                // Progressive send: small chunks so the client parser advances dynamically.
                for (int offset = 0; offset < sheetXml.Length; offset += ChunkSize) {
                    int len = Math.Min(ChunkSize, sheetXml.Length - offset);
                    serverNet.Write(sheetXml, offset, len);
                    serverNet.Flush();
                    Thread.Sleep(ChunkDelayMs);
                }
            }

            listener.Stop();
            clientTask.Wait(TimeSpan.FromSeconds(30));

            if (clientError != null) {
                Console.WriteLine("ERROR xml client: {0}", clientError.Message);
                return false;
            }
            if (rowsSeen <= 0) {
                Console.WriteLine("ERROR: no XmlRowPackage received");
                return false;
            }

            Console.WriteLine("-- XML OK (rows={0}, dimensions={1})", rowsSeen, dimSeen);
            return true;
        }

        /// <summary>
        /// Server sends length-prefixed BIFF record packages; client rebuilds BinaryPackage records live.
        /// </summary>
        private static bool RunBiffPackageDemo(string xlsPath) {
            Console.WriteLine("-- BIFF packages over TCP: {0}", Path.GetFileName(xlsPath));

            var listener = new TcpListener(IPAddress.Loopback, 0);
            listener.Start();
            int port = ((IPEndPoint)listener.LocalEndpoint).Port;

            int packagesSeen = 0;
            Exception clientError = null;

            var clientTask = Task.Run(() => {
                try {
                    using (var client = new TcpClient()) {
                        client.Connect(IPAddress.Loopback, port);
                        using (NetworkStream net = client.GetStream()) {
                            var recordReader = new ExcelBinaryReader();
                            var lenBuf = new byte[4];
                            while (ReadExact(net, lenBuf, 4)) {
                                int size = BitConverter.ToInt32(lenBuf, 0);
                                if (size <= 0) {
                                    break; // end marker
                                }
                                var body = new byte[size];
                                if (!ReadExact(net, body, size)) {
                                    throw new EndOfStreamException("truncated BIFF package");
                                }
                                XlsBiffRecord record = XlsBiffRecord.GetRecord(body, 0, recordReader);
                                if (record == null) {
                                    continue;
                                }
                                packagesSeen++;
                                // BinaryPackage = the record itself
                                BinaryPackage package = record;
                                if (packagesSeen <= 20 || packagesSeen % 200 == 0) {
                                    Console.WriteLine("[biff-stream] #{0} {1} size={2}",
                                        packagesSeen, package.GetType().Name, record.Size);
                                }
                            }
                        }
                    }
                } catch (Exception ex) {
                    clientError = ex;
                }
            });

            using (FileStream fs = File.Open(xlsPath, FileMode.Open, FileAccess.Read)) {
                var excelReader = new ExcelBinaryReader();
                excelReader.Open(fs);
                using (TcpClient serverClient = listener.AcceptTcpClient())
                using (NetworkStream serverNet = serverClient.GetStream()) {
                    foreach (BinaryPackage package in excelReader.StreamPackages()) {
                        var record = (XlsBiffRecord)package;
                        byte[] bytes = record.Bytes;
                        int size = record.Size;
                        byte[] len = BitConverter.GetBytes(size);
                        serverNet.Write(len, 0, 4);
                        serverNet.Write(bytes, 0, size);
                        serverNet.Flush();
                        Thread.Sleep(ChunkDelayMs);
                    }
                    // end marker
                    serverNet.Write(BitConverter.GetBytes(0), 0, 4);
                    serverNet.Flush();
                }
                excelReader.Close();
            }

            listener.Stop();
            clientTask.Wait(TimeSpan.FromSeconds(60));

            if (clientError != null) {
                Console.WriteLine("ERROR biff client: {0}", clientError.Message);
                return false;
            }
            if (packagesSeen <= 0) {
                Console.WriteLine("ERROR: no BIFF packages received");
                return false;
            }

            Console.WriteLine("-- BIFF OK (packages={0})", packagesSeen);
            return true;
        }

        private static byte[] LoadFirstWorksheetXml(string xlsxPath) {
            using (FileStream fs = File.Open(xlsxPath, FileMode.Open, FileAccess.Read)) {
                var zip = new ZipWorker();
                if (!zip.Open(fs, ownsStream: false)) {
                    return null;
                }
                try {
                    // Prefer sheet1; fall back to first worksheets/*.xml entry path used by test files.
                    Stream sheet = zip.GetWorksheetStream(1)
                                   ?? zip.GetWorksheetStream("worksheets/sheet1.xml")
                                   ?? zip.GetWorksheetStream("/xl/worksheets/sheet1.xml");
                    if (sheet == null) {
                        // Some workbooks start at sheet2 for the visible sheet — try a few ids.
                        for (int id = 2; id <= 5 && sheet == null; id++) {
                            sheet = zip.GetWorksheetStream(id);
                        }
                    }
                    if (sheet == null) {
                        return null;
                    }
                    using (sheet)
                    using (var ms = new MemoryStream()) {
                        sheet.CopyTo(ms);
                        return ms.ToArray();
                    }
                } finally {
                    zip.Dispose();
                }
            }
        }

        private static bool ReadExact(Stream stream, byte[] buffer, int count) {
            int offset = 0;
            while (offset < count) {
                int n = stream.Read(buffer, offset, count - offset);
                if (n <= 0) {
                    return false;
                }
                offset += n;
            }
            return true;
        }
    }
}
