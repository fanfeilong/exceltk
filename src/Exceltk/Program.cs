using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelToolKit {
    internal class Program {
        private static void Main(string[] args) {
            Xls2MarkDown(args);
        }

        private static void Xls2MarkDown(string[] args) {
            int ret=1;
            var cmd=new CommandParser(args);

            do {
                if (cmd["t"]!=null&&cmd["t"]=="md") {
                    if (cmd["xls"]==null) {
                        break;
                    }

                    string xls=cmd["xls"];
                    string sheet=cmd["sheet"];
                    string output=Path.Combine(Path.GetDirectoryName(xls), Path.GetFileNameWithoutExtension(xls));

                    if (!File.Exists(xls)) {
                        Console.WriteLine("xls file is not exist:{0}", xls);
                        break;
                    }

                    if (sheet!=null) {
                        Extension.MarkDownTable table=xls.ToMd(sheet);
                        string tableFile=output+table.Name+".md";
                        File.WriteAllText(tableFile, table.Value);
                        Console.WriteLine("Output File: {0}", tableFile);
                    } else {
                        IEnumerable<Extension.MarkDownTable> tables=xls.ToMd();
                        foreach (Extension.MarkDownTable table in tables) {
                            string tableFile=output+table.Name+".md";
                            File.WriteAllText(tableFile, table.Value);
                            Console.WriteLine("Output File: {0}", tableFile);
                        }
                    }

                    ret=0;
                }
            } while (false);

            if (ret!=0) {
                Console.WriteLine();
                Console.WriteLine("Usecase:exceltk -t md -xls xlsfile [-sheet sheetname]");
            }

            //Console.Read();
        }
    }
}