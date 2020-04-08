using CRSerializer;
using System;
using System.Configuration;
using System.Collections.Specialized;
using System.Diagnostics;
using System.IO;
using System.Reflection;

namespace CRDiff
{
    /* Features:
     * - Integratable with TortoiseGit (as well as subversion and stand-alone)
     * - Checks file extension, and if it is not .rpt, file passes directly to text diff tool.
     * - When a report, opens it/them. If error, output message
     * - Serialize all the report details into a text file we want to be able to compare, like
     *      Command query
     *      element dimensions and settings
     *      subreports
     * - Save serialized files and pass to the text diff tool
     * - Clean-up (delete) temp files after text diff tool is closed.
     *
     * Road Map:
     * - Add command-line switches such as /d[ifftool], /1[stReport], /2[ndReport], /t[empOutput]
     * - Add configuration (either in registry or app.config)
     * - Hook into shell context menu to select report files to diff.
     * - Serialization can be into either xml or json (command line switch -x[ml] or -j[son]
     *
     * Task list:
     * - Add versioning ie YY.MM.dd.gitCommit (done)
     *
       TODO: Need to serialize:
           element suppression formulas (and other formulas for elements)
           sub-textbox formatting (ie font changes, coloring, tab settings, etc)
           "Lock Position"
           "change number sign"
           Page Setup (ie Printer, margins)
           "Keep Data" setting

     * Testing arguments:
     * "C:\Program Files (x86)\Compare It!\wincmp3.exe" "C:\dev\svn\Reports\Crystal Reports\trunk\KinderMorgan_CS_MT_Testing.rpt" "C:\dev\svn\Reports\Crystal Reports\trunk\KinderMorgan_CS_MT.rpt"
     * "C:\dev\svn\Reports\Crystal Reports\trunk\KinderMorgan_CS_MT_Testing.rpt" KM_CS_MT_Testing.json
     */

    internal class CRDiff
    {
        private static void Main(string[] args)
        {
            switch (args.Length)
            {
                case 1:
                    {
                        // Assumption is a report file is being serialized to .json
                        if (Path.GetExtension(args[0]) == ".rpt")
                        {
                            SerializeToFile(args[0]);
                            //SerializeToStdOut(args[0]);
                        }
                        else
                        {
                            Console.WriteLine(string.Join(" ", args));
                            Console.WriteLine("Not a report file");
                            Usage();
                        }
                        break;
                    }
                case 2:
                    {
                        // Assumption is that the diff tool calls with the report file to convert
                        // and a temp filepath which it will clean up when done.
                        if (Path.GetExtension(args[0]) == ".rpt"
                            && File.Exists(args[0])
                            && Path.GetExtension(args[1]) != ".rpt" // 2nd param is not a report file
                            )
                        {
                                //Console.WriteLine($"Creating \"{args[1]}\" from \"{args[0]}\"");
                                SerializeToFile(args[0], args[1]);
                                return;
                        }

                        // If the 2nd filepath is also .rpt, check config for difftool and use it. If difftool doesn't exist, prompt.
                        // 2nd param is a report (presumably to compare with)
                        // Do we have a diff tool?
                        var diffTool = GetTextDiffTool();
                        if (string.IsNullOrEmpty(diffTool))
                        {
                            Console.WriteLine("Please provide the path to your text diff tool:");
                            diffTool = Console.ReadLine().Replace("\"","");
                        }

                        if (CompareFiles(diffTool, args[0], args[1]))
                        {
                            if (diffTool != GetTextDiffTool())
                            {
                                SetTextDiffTool(diffTool);
                            }
                        }
                        else
                        {
                            ReadKey();
                        }

                        break;
                    }
                case 3:
                    {
                        if(CompareFiles(args[0], args[1], args[2])
                            && args[0] != GetTextDiffTool())
                        {
                            SetTextDiffTool(args[0]);
                        }

                        break;
                    }
                default:
                    {
                        Usage();
                        ReadKey();
                        break;
                    }
            }
        } // end Main()

        /// <summary>
        /// Creates json files for two reports and sends them to the diff tool.
        /// If json files already exist, they are updated. If they didn't previously exist, they are deleted.
        /// </summary>
        /// <param name="diffTool"></param>
        /// <param name="filePath1"></param>
        /// <param name="filePath2"></param>
        /// <returns>true if successful</returns>
        private static bool CompareFiles(string diffTool, string filePath1, string filePath2)
        {
            string
                textFile1,
                textFile2;
            bool
                file1IsRpt = false,
                file2IsRpt = false,
                textFile1Exists = false,
                textFile2Exists = false;

            if (!ValidateFile(diffTool, "The diff tool")
                || !ValidateFile(filePath1, "The first file")
                || !ValidateFile(filePath2, "The second file"))
            {
                return false;
            }

            try
            {
                if (Path.GetExtension(filePath1) == ".rpt")
                {
                    file1IsRpt = true;
                    textFile1 = Path.ChangeExtension(filePath1, "json");
                    textFile1Exists = File.Exists(textFile1);
                    textFile1 = SerializeToFile(filePath1, textFile1, reportOrder: 1);
                }
                else
                {
                    textFile1 = filePath1;
                }

                if (Path.GetExtension(filePath2) == ".rpt")
                {
                    file2IsRpt = true;
                    textFile2 = Path.ChangeExtension(filePath2, "json");
                    textFile2Exists = File.Exists(textFile2);
                    textFile2 = SerializeToFile(filePath2, textFile2, reportOrder: 2);
                }
                else
                {
                    textFile2 = filePath2;
                }

                var diffParams = $"\"{textFile1}\" \"{textFile2}\"";
                Console.Write($"Starting \"{diffTool}\" {diffParams}");

                var diffProc = Process.Start(
                    diffTool,
                    diffParams
                    );

                diffProc.WaitForExit();

                // Clean up (delete) temp files after text diff tool is finished with them (if they didn't exist already)
                if (file1IsRpt && !textFile1Exists)
                    File.Delete(textFile1);

                if (file2IsRpt && !textFile2Exists)
                    File.Delete(textFile2);
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message} {ex.InnerException?.Message}");
                ReadKey();
                return false;
            }
        }

        private static bool ValidateFile(string filePath, string ifEmpty = "A file")
        {
            //TODO: If file is an executable, search path environment
            // ref: https://stackoverflow.com/questions/3855956/check-if-an-executable-exists-in-the-windows-path
            // ref: https://stackoverflow.com/a/58523041/2394826

            if (!File.Exists(filePath)) 
            {
                if (string.IsNullOrEmpty(filePath))
                {
                    filePath = ifEmpty;
                }
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"{filePath} was not found");
                Console.ResetColor();
                Usage();
                return false;
            }
            return true;
        }

        private static void Usage(string[] args = null)
        {
            Console.WriteLine("Usage:");
            Console.WriteLine("CRDiff ReportFilename.rpt");
            Console.WriteLine("     - Creates a serialized .json file of the report");
            Console.WriteLine("CRDiff ReportFilename.rpt TempFilename.json");
            Console.WriteLine("     - Writes json to TempFileName. Allows use as a .rpt converter in many text diff tools");
            Console.WriteLine("CRDiff ReportFilename1.rpt ReportFilename2.rpt");
            Console.WriteLine("CRDiff DiffAppPath.exe ReportFilename1.rpt ReportFilename2.rpt");
            Console.WriteLine("     - Serializes 2 reports and passes serialized .json files to DiffApp for comparison");
            Console.WriteLine();
            Console.WriteLine($"CRDiff Version: {CRDiffProductVersion()}");
            Console.WriteLine($"Path: {AppDomain.CurrentDomain.BaseDirectory}{Environment.GetCommandLineArgs()[0]}");
            Console.WriteLine();

            try
            {
                var version = (new CRSerialize()).CRVersion();
                Console.WriteLine($"CR Version: {version}");
            }
            catch (Exception e)
            {
                Console.WriteLine($"CR library is not available ({e.Message} {e.InnerException.Message})");
            }
        }

        private static void Instructions()
        {
            Console.WriteLine(@"

");
        }

        private static void ReadKey()
        {
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }

        private static string SerializeToFile(string rptFile, string textFile = null, int reportOrder = 1)
        {
            var serializer = new CRSerialize();

            textFile = textFile ?? Path.ChangeExtension(rptFile, "json");
            var file = File.CreateText(textFile);
            Console.WriteLine($"Loading {rptFile}");

            Console.WriteLine($"Saving {textFile}");
            var serialized = serializer.Serialize(rptFile, textFile, reportOrder);
            file.Write(serialized);
            file.Flush();
            file.Close();
            return textFile;
        }

        private static void SerializeToStdOut(string rptFile)
        {
            var serializer = new CRSerialize();
            Console.Write(serializer.Serialize(rptFile));
        }

        private static string CRDiffProductVersion()
        {
            return FileVersionInfo
                .GetVersionInfo(
                    Assembly
                    .GetExecutingAssembly()
                    .Location
                    )
                .ProductVersion;
        }

        internal static string GetAppSetting(string Key)
        {
            return ConfigurationManager.AppSettings.Get(Key);
        }
        internal static bool SetAppSetting(string Key, string Value)
        {
            bool result = false;
            try
            {
                var config = 
                  ConfigurationManager.OpenExeConfiguration(
                                       ConfigurationUserLevel.None);

                config.AppSettings.Settings.Remove(Key);
                var kvElem = new KeyValueConfigurationElement(Key, Value);
                config.AppSettings.Settings.Add(kvElem);

                // Save the configuration file.
                config.Save(ConfigurationSaveMode.Modified);

                // Force a reload of a changed section.
                ConfigurationManager.RefreshSection("appSettings");

                result = true;
            }
            finally
            { }
            return result;
        } // function

        private static string GetTextDiffTool()
        {
            return GetAppSetting("TextDiffTool");
        }
        private static bool SetTextDiffTool(string tool)
        {
            return SetAppSetting("TextDiffTool", tool);
        }
    }
}