using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

#if !OS_WINDOWS
using System.Xml;
#endif

using Exceltk.Format;
using Exceltk.Reader;

namespace Exceltk {
    public class SimpleTable {
        public string Name { get; set; }
        public string Value { get; set; }
    }

    public static class Extension {
        public static IEnumerable<SimpleTable> ToSimpleTable(this string xls, string ext) {
            IFormatPlugin plugin = FormatRegistry.GetRequired(ext);
            DataSet dataSet = WorkbookLoader.Load(xls);
            foreach (FormatArtifact artifact in plugin.Export(dataSet, null)) {
                if (artifact.IsBinary) {
                    continue;
                }
                yield return new SimpleTable {
                    Name = artifact.Name,
                    Value = artifact.TextContent
                };
            }
        }

        public static SimpleTable ToSimpleTable(this string xls, string sheet, string ext) {
            IFormatPlugin plugin = FormatRegistry.GetRequired(ext);
            DataSet dataSet = WorkbookLoader.Load(xls);
            foreach (FormatArtifact artifact in plugin.Export(dataSet, sheet)) {
                if (artifact.IsBinary) {
                    continue;
                }
                return new SimpleTable {
                    Name = artifact.Name,
                    Value = artifact.TextContent
                };
            }
            throw new ArgumentException("Sheet not found: " + sheet);
        }

        #if !OS_WINDOWS
        public static void Close(this XmlReader xmlReader) {
            xmlReader.Dispose();
        }
        public static void Close(this Stream stream) {
            stream.Dispose();
        }
        public static void Close(this MemoryStream stream) {
            stream.Dispose();
        }
        public static void Close(this BinaryReader stream) {
            stream.Dispose();
        }
        #endif

        public static string GetUserName() {
            #if OS_WINDOWS
                return Environment.UserName;
            #else
                return "";
            #endif
        }

        public static Encoding DefaultEncoding() {
            #if OS_WINDOWS
                return Encoding.Default;
            #else
                return Encoding.UTF8;
            #endif
        }

        public static bool IsHexDigit(this char digit) {
            return (('0' <= digit && digit <= '9') ||
                    ('a' <= digit && digit <= 'f') ||
                    ('A' <= digit && digit <= 'F'));
        }
        public static int FromHex(this char digit) {
            if ('0' <= digit && digit <= '9') {
                return (int)(digit - '0');
            }

            if ('a' <= digit && digit <= 'f')
                return (int)(digit - 'a' + 10);

            if ('A' <= digit && digit <= 'F')
                return (int)(digit - 'A' + 10);

            throw new ArgumentException("digit");
        }
    }
}
