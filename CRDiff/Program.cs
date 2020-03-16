using CRSerializer;
//using CrystalDecisions.CrystalReports.Engine;
using System;
//using System.Collections.Generic;
using System.Diagnostics;
using System.IO;

namespace CRDiff
{
    /* Features:
     * - Integratable with Github (as well as subversion and stand-alone)
     * - Checks file extension, and if it is not .rpt, file passes directly to text diff tool.
     * - When a report, opens it/them. If error, output message
     * - Serialize all the report details into a text file we want to be able to compare, like
     *      Command query
     *      element dimensions and settings
     *      subreports
     * - Serialization can be into either xml or json (command line switch -x[ml] or -j[son]
     * - Save serialized files and pass to the text diff tool
     * - Clean-up (delete) temp files after text diff tool is closed.
     *
     * Road Map:
     * - Hook into shell context menu to select report files to diff.
     *
     * Task list:
     * -
       TODO: Need to show:
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
                            && File.Exists(args[0]))
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
                            file2IsRpt = false;

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
                                textFile1 = SerializeToFile(rptFile1, reportOrder: 1);
                            }
                            else
                            {
                                textFile1 = rptFile1;
                            }

                            if (Path.GetExtension(rptFile2) == ".rpt")
                            {
                                file2IsRpt = true;
                                textFile2 = SerializeToFile(rptFile2, reportOrder: 2);
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

                            // Clean up (delete) temp files after text diff tool is finished with them
                            if (file1IsRpt)
                                File.Delete(textFile1);

                            if (file2IsRpt)
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
            Console.WriteLine("CRDiff PathToReport (Create a serialized .json file of the report)");
            Console.WriteLine("CRDiff PathToReport TargetFilename (Use CRDiff as a converter in many text diff tools)");
            Console.WriteLine("CRDiff PathToTextDiffApp PathToReport1 PathToReport2 (Configure SC Diff Viewer for .rpt files)");
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

        private static string SerializeToFile(string rptPath, string textPath = null, int reportOrder = 1)
        {

            var serializer = new CRSerialize();

            textPath = textPath ?? Path.ChangeExtension(rptPath, "json");
            var file = File.CreateText(textPath);
            Console.WriteLine($"Loading {rptPath}");

            Console.WriteLine($"Saving {textPath}");
            var serialized = serializer.Serialize(rptPath, textPath, reportOrder);
            file.Write(serialized);
            file.Flush();
            file.Close();
            return textPath;
        }

    }
}