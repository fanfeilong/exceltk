using System;
using System.Collections.Generic;
using System.IO;

using Exceltk.Util;

namespace Exceltk {
    internal class Program {

        [STAThread]
        private static void Main(string[] args) {
            var cmd=new CommandParser(args);
            var r = InitConfig(cmd);
            if (r) {
                Xls2MarkDown(cmd);
            }
        }

        private static bool InitConfig(CommandParser cmd) {
            // default
            Config.DecimalPrecision = 0;
            if (cmd["t"]!=null){
                if (cmd["t"] == "md"){
                    if (cmd["bhead"]!=null) {
                        Config.BodyHead = true;
                    } else {
                        Config.BodyHead = false;
                    }

                    if (cmd["p"]!=null) {
                        int precision = 0;
                        var ret = Int32.TryParse(cmd["p"],out precision);
                        if (ret) {
                            if (precision > 10) {
                                Console.WriteLine("presision too larger:" + precision);
                                return false;
                            }
                            if (precision > 0) {
                                Config.DecimalPrecision = precision;
                            }
                        }
                    }

                    if(cmd["a"]!=null){
                        var align = cmd["a"];
                        Config.TableAligin = align;
                    }else{
                        Config.TableAligin = "l";
                    }
                }

                if (cmd["t"] == "tex") {
                    if (cmd["sn"] != null) {
                        Config.SplitNumber = true;
                    } else {
                        Config.SplitNumber = false;
                    }

                    if (cmd["st"] != null) {
                        Config.SplitTable = true;
                        int rows = 0;
                        bool ret = Int32.TryParse(cmd["st"], out rows);
                        if (ret) {
                            Config.SplitTableRow = rows;
                        } else {
                            Config.SplitTable = false;
                        }
                    } else {
                        Config.SplitTable = false;
                    }
                }
            }
            return true;
        }

        private static void Xls2MarkDown(CommandParser cmd) {
            int ret=1;
            do {
                // check target
                if (cmd["t"]==null) {
                    Console.WriteLine("ERROR:target not found");
                    break;
                }

                // check xls arg
                if (cmd["xls"] == null) {
                    Console.WriteLine("ERROR:xls not found");
                    break;
                }

                // check xls exist
                string xls = cmd["xls"];
                string sheet = cmd["sheet"];
                string root = Directory.GetCurrentDirectory();
                xls = Path.Combine(root, xls);
                if (!File.Exists(xls)) {
                    Console.WriteLine("ERROR:xls file is not exist:{0}", xls);
                    break;
                }

                // check xls path
                string dirName = Path.GetDirectoryName(xls);
                string fileName = Path.GetFileNameWithoutExtension(xls);
                if(dirName==null||fileName==null){
                    Console.WriteLine("ERROR: xls path is valid:{0}",xls);
                    break;
                }

                // check md or json target
                var target = cmd["t"];
                if(target!="md"&&target!="json"&&target!="tex"){
                    Console.WriteLine("ERROR: target not support",target);
                    break;
                }

                // output
                var output=Path.Combine(dirName, fileName);
                if (sheet!=null) {
                    var table=xls.ToSimpleTable(sheet, target);
                    string tableFile=output+table.Name+"."+ target;
                    File.WriteAllText(tableFile, table.Value);
                    Console.WriteLine("Output File: {0}", tableFile);
                } else {
                    var tables=xls.ToSimpleTable(target);
                    foreach (var table in tables) {
                        string tableFile=output+table.Name+"."+ target;
                        File.WriteAllText(tableFile, table.Value);
                        Console.WriteLine("Output File: {0}", tableFile);
                    }
                }
                ret=0;
                Console.WriteLine("Done!");


            } while (false);

            if (ret!=0) {
                Console.WriteLine();
                Console.WriteLine("Usecase:");
                Console.WriteLine("1. Convert xls to markdown: Exceltk -t md -xls xlsfile [-sheet sheetname]");
                Console.WriteLine("2. Monitor and convert clipboard to markdown: Exceltk -t cm");
            }
        }
    }
}