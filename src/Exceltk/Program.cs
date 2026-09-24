using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

using Exceltk.Format;
using Exceltk.Reader;
using Exceltk.Util;

namespace Exceltk {
    internal class Program {

        [STAThread]
        private static void Main(string[] args) {
            var cmd = new CommandParser(args);
            if (cmd["t"] == "tcpstream") {
                string root = Directory.GetCurrentDirectory();
                string xlsx = cmd["xlsx"] ?? cmd["xls"] ?? Path.Combine("test", "test1.xlsx");
                string xls = cmd["biff"] ?? Path.Combine("test", "test8.xls");
                if (!Path.IsPathRooted(xlsx)) {
                    xlsx = Path.Combine(root, xlsx);
                }
                if (!Path.IsPathRooted(xls)) {
                    xls = Path.Combine(root, xls);
                }
                Environment.ExitCode = Test.TcpStreamingDemo.Run(xlsx, xls);
                return;
            }

            var r = InitConfig(cmd);
            if (r) {
                RunConversion(cmd);
            }
        }

        private static bool InitConfig(CommandParser cmd) {
            // default
            Config.DecimalPrecision = 0;
            if (cmd["t"] != null) {
                if (cmd["t"] == "md") {
                    if (cmd["bhead"] != null) {
                        Config.BodyHead = true;
                    } else {
                        Config.BodyHead = false;
                    }

                    Config.PrettyTable = cmd["pretty"] != null;
                    Config.MultiMarkdown = cmd["mmd"] != null;

                    if (cmd["p"] != null) {
                        int precision = 0;
                        var ret = Int32.TryParse(cmd["p"], out precision);
                        if (ret) {
                            if (precision > 10) {
                                Console.WriteLine("presision too larger:" + precision);
                                return false;
                            }
                            if (precision >= 0) {
                                Config.DecimalPrecision = precision;
                                Config.HasDecimalPrecision = true;
                            }
                        }
                    }

                    if (cmd["a"] != null) {
                        var align = cmd["a"];
                        Config.TableAligin = align;
                    } else {
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

        private static void RunConversion(CommandParser cmd) {
            int ret = 1;
            do {
                if (cmd["t"] == null) {
                    Console.WriteLine("ERROR:target not found");
                    break;
                }
                var target = cmd["t"];
                if (target == "cm") {
                    Console.WriteLine("ERROR: -t cm (clipboard monitor GUI) was removed after version 0.0.9.");
                    Console.WriteLine("       Use 0.0.9 on Windows if you need that feature:");
                    Console.WriteLine("       http://files.cnblogs.com/files/math/exceltk0.0.9.7z");
                    break;
                }

                IFormatPlugin plugin;
                if (!FormatRegistry.TryGet(target, out plugin)) {
                    Console.WriteLine("ERROR: target not support: {0}", target);
                    Console.WriteLine("Supported plugins: {0}", string.Join(", ", PluginNames()));
                    break;
                }

                bool importMode = cmd["import"] != null;
                if (importMode) {
                    if (!plugin.SupportsImport) {
                        Console.WriteLine("ERROR: format plugin '{0}' does not support import", plugin.Name);
                        break;
                    }
                    ret = RunImport(cmd, plugin);
                    break;
                }

                if (!plugin.SupportsExport) {
                    Console.WriteLine("ERROR: format plugin '{0}' does not support export", plugin.Name);
                    break;
                }
                ret = RunExport(cmd, plugin);
            } while (false);

            if (ret != 0) {
                PrintUsage();
            }
            Environment.ExitCode = ret;
        }

        private static int RunExport(CommandParser cmd, IFormatPlugin plugin) {
            if (cmd["xls"] == null && cmd["csv"] == null) {
                Console.WriteLine("ERROR:xls/csv not found");
                return 1;
            }

            string xls = cmd["xls"] ?? cmd["csv"];
            string sheet = cmd["sheet"];
            string root = Directory.GetCurrentDirectory();
            if (!Path.IsPathRooted(xls)) {
                xls = Path.Combine(root, xls);
            }
            if (!File.Exists(xls)) {
                Console.WriteLine("ERROR:input file is not exist:{0}", xls);
                return 1;
            }
            if (!WorkbookLoader.IsSupportedExtension(xls)) {
                Console.WriteLine("ERROR:unsupported file format:{0}", Path.GetExtension(xls));
                return 1;
            }

            string dirName = Path.GetDirectoryName(xls);
            string fileName = Path.GetFileNameWithoutExtension(xls);
            if (dirName == null || fileName == null) {
                Console.WriteLine("ERROR: xls path is valid:{0}", xls);
                return 1;
            }

            string outputBase = cmd["out"] != null
                ? (Path.IsPathRooted(cmd["out"]) ? cmd["out"] : Path.Combine(root, cmd["out"]))
                : Path.Combine(dirName, fileName);

            DataSet dataSet = WorkbookLoader.Load(xls);
            IEnumerable<FormatArtifact> artifacts = plugin.Export(dataSet, sheet);
            int written = 0;
            foreach (FormatArtifact artifact in artifacts) {
                string tableFile = outputBase + artifact.Name + "." + artifact.Extension;
                string outDir = Path.GetDirectoryName(tableFile);
                if (!string.IsNullOrEmpty(outDir) && !Directory.Exists(outDir)) {
                    Directory.CreateDirectory(outDir);
                }
                if (artifact.IsBinary) {
                    File.WriteAllBytes(tableFile, artifact.BinaryContent);
                } else {
                    // UTF-8 without BOM so JSON/md/tex stay machine-friendly.
                    File.WriteAllText(tableFile, artifact.TextContent, new UTF8Encoding(false));
                }
                Console.WriteLine("Output File: {0}", tableFile);
                written++;
            }
            if (written == 0) {
                Console.WriteLine("ERROR: no sheets exported");
                return 1;
            }
            Console.WriteLine("Done!");
            return 0;
        }

        private static int RunImport(CommandParser cmd, IFormatPlugin plugin) {
            string input = cmd["import"];
            if (string.IsNullOrEmpty(input) || input == "true") {
                // Allow: -import file.ext   (CommandParser stores value after flag)
                // If flag was bare, try -xls as the import source.
                input = cmd["xls"] ?? cmd["csv"];
            }
            if (string.IsNullOrEmpty(input) || input == "true") {
                Console.WriteLine("ERROR:import input not found (use -import path)");
                return 1;
            }

            string root = Directory.GetCurrentDirectory();
            if (!Path.IsPathRooted(input)) {
                input = Path.Combine(root, input);
            }
            if (!File.Exists(input)) {
                Console.WriteLine("ERROR:import file is not exist:{0}", input);
                return 1;
            }

            string outPath = cmd["out"];
            if (string.IsNullOrEmpty(outPath) || outPath == "true") {
                string dir = Path.GetDirectoryName(input) ?? root;
                string name = Path.GetFileNameWithoutExtension(input);
                outPath = Path.Combine(dir, name + ".xlsx");
            } else if (!Path.IsPathRooted(outPath)) {
                outPath = Path.Combine(root, outPath);
            }
            if (!outPath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)) {
                outPath = outPath + ".xlsx";
            }

            try {
                using (FileStream stream = File.OpenRead(input)) {
                    DataSet dataSet = plugin.Import(stream, input);
                    XlsxWorkbookWriter.Write(dataSet, outPath);
                }
            } catch (Exception ex) {
                Console.WriteLine("ERROR:import failed: {0}", ex.Message);
                return 1;
            }

            Console.WriteLine("Output File: {0}", outPath);
            Console.WriteLine("Done!");
            return 0;
        }

        private static IEnumerable<string> PluginNames() {
            foreach (IFormatPlugin plugin in FormatRegistry.All) {
                yield return plugin.Name;
            }
        }

        private static void PrintUsage() {
            Console.WriteLine();
            Console.WriteLine("Usecase:");
            Console.WriteLine("1. Export xls/xlsx/csv via format plugin: Exceltk -t md|json|tex|img -xls file [-sheet name]");
            Console.WriteLine("2. Import format file back to Excel: Exceltk -t md|json|tex|img -import file [-out out.xlsx]");
            Console.WriteLine("3. Pretty markdown: Exceltk -t md -pretty -xls file");
            Console.WriteLine("4. MultiMarkdown HTML tables: Exceltk -t md -mmd -xls file");
            Console.WriteLine("5. Marked PNG only: Exceltk -t img -xls file.xlsx  /  Exceltk -t img -import file.png");
            Console.WriteLine("   (unmarked images are rejected)");
            Console.WriteLine("Note: -t cm (clipboard GUI) was removed after 0.0.9; use 0.0.9 if needed.");
            Console.WriteLine("6. TCP streaming package demo: Exceltk -t tcpstream [-xlsx file.xlsx] [-biff file.xls]");
            Console.WriteLine();
            Console.WriteLine("Format plugins:");
            foreach (IFormatPlugin plugin in FormatRegistry.All) {
                Console.WriteLine("  -t {0}  ({1})  export={2} import={3}",
                    plugin.Name, plugin.Description, plugin.SupportsExport, plugin.SupportsImport);
            }
        }
    }
}
