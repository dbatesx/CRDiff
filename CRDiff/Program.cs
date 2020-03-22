using CRSerializer;
using System;
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

    internal class Program
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
                            && Path.GetExtension(args[1]) != ".rpt"
                            )
                        {
                            //Console.WriteLine($"Creating \"{args[1]}\" from \"{args[0]}\"");
                            SerializeToFile(args[0], args[1]);
                        }
                        else
                        {
                            Console.WriteLine($"Create \"{args[1]}\" from \"{args[0]}\"?");
                            Console.WriteLine(string.Join(" ", args));
                            Usage();
                        }
                        break;
                    }
                case 3:
                    {
                        string
                            textDiffTool = args[0],
                            rptFile1 = args[1],
                            rptFile2 = args[2],
                            textFile1,
                            textFile2;
                        bool
                            file1IsRpt = false,
                            file2IsRpt = false,
                            textFile1Exists = false,
                            textFile2Exists = false;

                        if (!File.Exists(textDiffTool)
                            || !File.Exists(rptFile1)
                            || !File.Exists(rptFile2))
                        {
                            Console.WriteLine(string.Join(" ", args));
                            Usage();
                            return;
                        }

                        try
                        {
                            if (Path.GetExtension(rptFile1) == ".rpt")
                            {
                                file1IsRpt = true;
                                textFile1 = Path.ChangeExtension(rptFile1, "json");
                                textFile1Exists = File.Exists(textFile1);
                                textFile1 = SerializeToFile(rptFile1, textFile1, reportOrder: 1);
                            }
                            else
                            {
                                textFile1 = rptFile1;
                            }

                            if (Path.GetExtension(rptFile2) == ".rpt")
                            {
                                file2IsRpt = true;
                                textFile2 = Path.ChangeExtension(rptFile2, "json");
                                textFile2Exists = File.Exists(textFile2);
                                textFile2 = SerializeToFile(rptFile2, textFile2, reportOrder: 2);
                            }
                            else
                            {
                                textFile2 = rptFile2;
                            }

                            var diffParams = $"\"{textFile1}\" \"{textFile2}\"";
                            Console.Write($"Starting \"{ textDiffTool}\" {diffParams}");

                            var diffProc = Process.Start(
                                textDiffTool,
                                diffParams
                                );

                            diffProc.WaitForExit();

                            // Clean up (delete) temp files after text diff tool is finished with them (if they didn't exist already)
                            if (file1IsRpt && !textFile1Exists)
                                File.Delete(textFile1);

                            if (file2IsRpt && !textFile2Exists)
                                File.Delete(textFile2);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Error: {ex.Message} {ex.InnerException?.Message}");
                            ReadKey();
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
        }

        private static void Usage(string[] args = null)
        {
            Console.WriteLine("Usage:");
            Console.WriteLine("CRDiff ReportFilename");
            Console.WriteLine("     - Creates a serialized .json file of the report");
            Console.WriteLine("CRDiff ReportFilename TempFilename");
            Console.WriteLine("     - Writes json to TempFileName. Allows use as a .rpt converter in many text diff tools");
            Console.WriteLine("CRDiff DiffAppPath ReportFilename1 ReportFilename2");
            Console.WriteLine("     - Serializes 2 reports and passes serialized .json files to DiffApp for comparison");
            Console.WriteLine();
            Console.WriteLine($"CRDiff Version: {CRDiffProductVersion()}");
            Console.WriteLine($"Path: {Environment.GetCommandLineArgs()[0]}");
            Console.WriteLine($"Directory: {AppDomain.CurrentDomain.BaseDirectory}");
            Console.WriteLine();

            try
            {
                var version = (new CRSerialize()).CRVersion();
                Console.WriteLine($"CR Version: {version}");
            }
            catch (Exception e)
            {
                Console.WriteLine($"CR library is not available ({e.Message})");
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
    }
}