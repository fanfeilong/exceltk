using System;
using System.Collections;
using System.Diagnostics;
using System.IO;
using ICSharpCode.SharpZipLib.Zip;

namespace Exceltk.Reader {
    /// <summary>
    /// OpenXML ZIP IO: keeps the archive open and streams individual entries
    /// (no full extract-to-temp). Worksheet XML can then be package-parsed
    /// incrementally as the entry inflate stream produces tokens.
    /// </summary>
    public class ZipWorker : IDisposable {
        private const string FILE_sharedStrings = "xl/sharedStrings.xml";
        private const string FILE_styles = "xl/styles.xml";
        private const string FILE_workbook = "xl/workbook.xml";
        private const string FILE_workbook_rels = "xl/_rels/workbook.xml.rels";

        private ZipFile m_zip;
        private Stream m_zipStream;
        private bool m_ownsStream;
        private string m_exceptionMessage;
        private bool m_isValid;
        private bool disposed;

        public bool IsValid {
            get {
                return m_isValid;
            }
        }

        public string ExceptionMessage {
            get {
                return m_exceptionMessage;
            }
        }

        public void Dispose() {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Open an .xlsx archive for streamed entry access (preferred over full extract).
        /// </summary>
        public bool Open(Stream fileStream) {
            return Open(fileStream, ownsStream: true);
        }

        public bool Open(Stream fileStream, bool ownsStream) {
            if (fileStream == null) {
                return false;
            }

            CloseZip();
            m_ownsStream = ownsStream;
            m_zipStream = fileStream;
            m_isValid = true;
            m_exceptionMessage = null;

            try {
                m_zip = new ZipFile(fileStream);
                m_zip.IsStreamOwner = false; // ZipWorker owns the stream lifetime
                if (FindEntry(FILE_workbook) == null) {
                    m_isValid = false;
                    m_exceptionMessage = "Missing xl/workbook.xml";
                    CloseZip();
                    return false;
                }
            } catch (Exception ex) {
                m_isValid = false;
                m_exceptionMessage = ex.Message;
                CloseZip();
                return false;
            }

            return m_isValid;
        }

        /// <summary>Backward-compatible alias for <see cref="Open(Stream)"/>.</summary>
        public bool Extract(Stream fileStream) {
            return Open(fileStream);
        }

        public Stream GetSharedStringsStream() {
            return OpenEntry(FILE_sharedStrings);
        }

        public Stream GetStylesStream() {
            return OpenEntry(FILE_styles);
        }

        public Stream GetWorkbookStream() {
            return OpenEntry(FILE_workbook);
        }

        public Stream GetWorkbookRelsStream() {
            return OpenEntry(FILE_workbook_rels);
        }

        public Stream GetWorksheetStream(int sheetId) {
            return OpenEntry(string.Format("xl/worksheets/sheet{0}.xml", sheetId));
        }

        public Stream GetWorksheetStream(string sheetPath) {
            return OpenEntry(NormalizeXlPath(sheetPath));
        }

        public Stream GetWorksheetRelsStream(string sheetPath) {
            string path = NormalizeXlPath(sheetPath);
            int slash = path.LastIndexOf('/');
            if (slash < 0) {
                return null;
            }
            string dir = path.Substring(0, slash);
            string file = path.Substring(slash + 1);
            return OpenEntry(dir + "/_rels/" + file + ".rels");
        }

        private static string NormalizeXlPath(string sheetPath) {
            if (string.IsNullOrEmpty(sheetPath)) {
                return sheetPath;
            }
            sheetPath = sheetPath.Replace('\\', '/');
            if (sheetPath.StartsWith("/")) {
                sheetPath = sheetPath.Substring(1);
            }
            if (sheetPath.StartsWith("xl/")) {
                return sheetPath;
            }
            return "xl/" + sheetPath;
        }

        private Stream OpenEntry(string entryName) {
            if (m_zip == null || string.IsNullOrEmpty(entryName)) {
                return null;
            }
            ZipEntry entry = FindEntry(entryName);
            if (entry == null || !entry.IsFile) {
                return null;
            }
            return m_zip.GetInputStream(entry);
        }

        private ZipEntry FindEntry(string entryName) {
            entryName = entryName.Replace('\\', '/');
            // ZipFile indexer / GetEntry expects the stored name
            IEnumerator enumerator = m_zip.GetEnumerator();
            while (enumerator.MoveNext()) {
                var entry = (ZipEntry)enumerator.Current;
                if (entry == null || string.IsNullOrEmpty(entry.Name)) {
                    continue;
                }
                string name = entry.Name.Replace('\\', '/');
                if (string.Equals(name, entryName, StringComparison.OrdinalIgnoreCase)) {
                    return entry;
                }
            }
            return null;
        }

        private void CloseZip() {
            if (m_zip != null) {
                m_zip.Close();
                m_zip = null;
            }
            if (m_ownsStream && m_zipStream != null) {
                m_zipStream.Dispose();
                m_zipStream = null;
            } else {
                m_zipStream = null;
            }
        }

        private void Dispose(bool disposing) {
            if (!disposed) {
                if (disposing) {
                    CloseZip();
                }
                disposed = true;
            }
        }

        ~ZipWorker() {
            Dispose(false);
        }
    }
}
