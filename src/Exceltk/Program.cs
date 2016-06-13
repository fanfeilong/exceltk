using exceltk.Clipborad;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ExcelToolKit {
    internal class Program {

        [DllImport("kernel32.dll", SetLastError = true)]
        static extern bool AllocConsole();

        [DllImport("kernel32.dll", SetLastError = true)]
        static extern bool FreeConsole();

        [DllImport("kernel32", SetLastError = true)]
        static extern bool AttachConsole(int dwProcessId);

        [STAThread]
        private static void Main(string[] args) {
            Xls2MarkDown(args);
            return;
        }

        private static void Xls2MarkDown(string[] args) {
            int ret=1;
            var cmd=new CommandParser(args);
            var allocConsole = false;
            do {
                if (cmd["t"]!=null) {
                    if (cmd["t"] == "md") {
                        if (!AttachConsole(-1)) {
                            AllocConsole();
                            allocConsole = true;
                        }
                        
                        if (cmd["xls"] == null) {
                            break;
                        }

                        string xls = cmd["xls"];
                        string sheet = cmd["sheet"];
                        string output = Path.Combine(Path.GetDirectoryName(xls), Path.GetFileNameWithoutExtension(xls));

                        if (!File.Exists(xls)) {
                            Console.WriteLine("xls file is not exist:{0}", xls);
                            break;
                        }

                        if (sheet != null) {
                            Extension.MarkDownTable table = xls.ToMd(sheet);
                            string tableFile = output + table.Name + ".md";
                            File.WriteAllText(tableFile, table.Value);
                            Console.WriteLine("Output File: {0}", tableFile);
                        } else {
                            IEnumerable<Extension.MarkDownTable> tables = xls.ToMd();
                            foreach (Extension.MarkDownTable table in tables) {
                                string tableFile = output + table.Name + ".md";
                                File.WriteAllText(tableFile, table.Value);
                                Console.WriteLine("Output File: {0}", tableFile);
                            }
                        }
                        ret = 0;
                        Console.WriteLine("Done!");
                    } else if (cmd["t"] == "cm") {
                        Application.Run(new ClipboradMonitor());
                        ret = 0;
                    } else {
                        // Ignore
                    }
                }

            } while (false);

            
            

            if (ret!=0) {
                Console.WriteLine();
                Console.WriteLine("Usecase:");
                Console.WriteLine("1. Convert xls to markdown: exceltk -t md -xls xlsfile [-sheet sheetname]");
                Console.WriteLine("2. Monitor and convert clipboard to markdown: exceltk -t cm");
            }

            SendKeys.SendWait("{ENTER}");
            if (allocConsole) {
                FreeConsole();
            }

            System.Environment.Exit(0);
            Application.ExitThread();
        }
    }
}