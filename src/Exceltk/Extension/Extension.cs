using System;
using System.Text;
using System.Collections.Generic;

#if !OS_WINDOWS
using System.Xml;
using System.IO;
#endif

namespace Exceltk {
    public class SimpleTable{
        public string Name{get;set;}
        public string Value{get;set;}
    }

    public static class Extension {
        public static IEnumerable<SimpleTable> ToSimpleTable(this string xls,string ext){
            if(ext=="md"){
                return xls.ToMd();
            }else if(ext=="json"){
                return xls.ToJson();
            } else if (ext == "tex") {
                return xls.ToTex();
            } else {
                return null;
            }
        }

        public static SimpleTable ToSimpleTable(this string xls,string sheet,string ext){
            if(ext=="md"){
                return xls.ToMd(sheet);
            }else if(ext=="json"){
                return xls.ToJson(sheet);
            } else if (ext == "tex") {
                return xls.ToTex(sheet);
            } else {
                return null;
            }
        }

        #if !OS_WINDOWS
        public static void Close(this XmlReader xmlReader) {
            xmlReader.Dispose();
        }
        public static void Close(this Stream stream){
            stream.Dispose();
        }
        public static void Close(this MemoryStream stream){
            stream.Dispose();
        }
        public static void Close(this BinaryReader stream){
            stream.Dispose();
        }
        #endif

        public static string GetUserName(){
            #if OS_WINDOWS
                return Environment.UserName;
            #else
                return "";
            #endif
        }

        public static Encoding DefaultEncoding(){
            #if OS_WINDOWS
                return Encoding.Default;
            #else
                return Encoding.UTF8;
            #endif
        }

        public static bool IsHexDigit(this char digit) {
            return (('0'<=digit&&digit<='9')||
                    ('a'<=digit&&digit<='f')||
                    ('A'<=digit&&digit<='F'));
        }
        public static int FromHex(this char digit) {
            if ('0'<=digit&&digit<='9') {
                return (int)(digit-'0');
            }

            if ('a'<=digit&&digit<='f')
                return (int)(digit-'a'+10);

            if ('A'<=digit&&digit<='F')
                return (int)(digit-'A'+10);

            throw new ArgumentException("digit");
        }
    }
}