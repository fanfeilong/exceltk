using System;
using System.Collections.Generic;
using System.IO;

#if OS_WINDOWS
using exceltk.Clipborad;
using System.Runtime.InteropServices;
using System.Windows.Forms;
#endif

namespace ExcelToolKit {
    internal class Program {
        
        #if OS_WINDOWS
        [DllImport("kernel32.dll", SetLastError = true)]
        static extern bool AllocConsole();

        [DllImport("kernel32.dll", SetLastError = true)]
        static extern bool FreeConsole();

        [DllImport("kernel32", SetLastError = true)]
        static extern bool AttachConsole(int dwProcessId);
        #endif

        [STAThread]
        private static void Main(string[] args) {
            var cmd=new CommandParser(args);
            InitConfig(cmd);
            Xls2MarkDown(cmd);
            return;
        }

        private static void InitConfig(CommandParser cmd) {
            // default
            Config.DecimalPrecision = 0;
            if (cmd["t"]!=null){
                if (cmd["t"] == "md"){
                    if (cmd["p"]!=null) {
                        int precision = 0;
                        Int32.TryParse(cmd["p"],out precision);
                        if(precision>0){
                            Config.DecimalPrecision=precision;                            
                        }
                    }
                }
            }
        }

        private static void Xls2MarkDown(CommandParser cmd) {
            int ret=1;
            #if OS_WINDOWS
            var allocConsole = false;
            #endif
            do {
                if (cmd["t"]!=null) {
                    if (cmd["t"] == "md") {
                        #if OS_WINDOWS
                        if (!AttachConsole(-1)) {
                            AllocConsole();
                            allocConsole = true;
                        }
                        #endif

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
                        #if OS_WINDOWS
                        Application.Run(new ClipboradMonitor());
                        #endif
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

            #if OS_WINDOWS
            SendKeys.SendWait("{ENTER}");
            if (allocConsole) {
                FreeConsole();
            }

            System.Environment.Exit(0);
            Application.ExitThread();
            #endif
        }
    }
}