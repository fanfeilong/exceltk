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
        /// Server streams worksheet XML bytes; client XmlPackageParser Pump/PopPackage
        /// emits Office element packages (dimension/row/…) as each completes.
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
                                while (parser.TryFrameOne() || parser.HavePackage) {
                                    while (parser.HavePackage) {
                                        XmlPackage package = parser.PopPackage();
                                        if (package is XmlDimensionPackage dim) {
                                            dimSeen++;
                                            Console.WriteLine("[xml-stream] dimension ref={0}", dim.Ref);
                                            if (parser.NamespaceUri != null) {
                                                Console.WriteLine("[xml-stream] worksheet ns={0}",
                                                    parser.NamespaceUri);
                                            }
                                        } else if (package is XmlRowPackage row) {
                                            rowsSeen++;
                                            if (rowsSeen <= 15 || rowsSeen % 25 == 0) {
                                                Console.WriteLine("[xml-stream] row#{0} cells={1} r={2}",
                                                    rowsSeen, row.Cells.Count, row.RowIndexAttribute);
                                            }
                                        } else if (package is XmlMergePackage merge) {
                                            Console.WriteLine("[xml-stream] mergeCell ref={0}", merge.Ref);
                                        } else if (package is XmlHyperlinkPackage link) {
                                            Console.WriteLine("[xml-stream] hyperlink ref={0}", link.Ref);
                                        }
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
        /// Server streams raw workbook BIFF bytes; client <see cref="BinaryPackageParser"/>
        /// PushData-frames records and pops packages as each header+body completes.
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
                            var parser = new BinaryPackageParser(recordReader);
                            var buf = new byte[ChunkSize];
                            int n;
                            while ((n = net.Read(buf, 0, buf.Length)) > 0) {
                                parser.PushData(buf, 0, n);
                                while (parser.HavePackage) {
                                    BinaryPackage package = parser.PopPackage();
                                    packagesSeen++;
                                    if (packagesSeen <= 20 || packagesSeen % 200 == 0) {
                                        Console.WriteLine("[biff-stream] #{0} {1} size={2}",
                                            packagesSeen, package.GetType().Name, package.Size);
                                    }
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
                byte[] workbookBytes = excelReader.WorkbookBytes;
                using (TcpClient serverClient = listener.AcceptTcpClient())
                using (NetworkStream serverNet = serverClient.GetStream()) {
                    for (int offset = 0; offset < workbookBytes.Length; offset += ChunkSize) {
                        int len = Math.Min(ChunkSize, workbookBytes.Length - offset);
                        serverNet.Write(workbookBytes, offset, len);
                        serverNet.Flush();
                        Thread.Sleep(ChunkDelayMs);
                    }
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
                    Stream sheet = zip.GetWorksheetStream(1)
                                   ?? zip.GetWorksheetStream("worksheets/sheet1.xml")
                                   ?? zip.GetWorksheetStream("/xl/worksheets/sheet1.xml");
                    if (sheet == null) {
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
    }
}
