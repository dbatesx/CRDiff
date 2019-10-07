using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CommandLine;
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;
//using System.Runtime.Serialization.json;
using CrystalDecisions.CrystalReports.Engine;
using System.Reflection;
using System.IO;
using System.Diagnostics;

namespace CRDiff
{
    /* Features:
     * - Integratable with Github (as well as subversion and stand-alone)
     * - Checks file extension, and if not .rpt passes directly to text diff tool.
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
     */
    class Program
    {
        //public class Options
        //{
        //    [Option('v', "verbose", Required = false, HelpText = "Set output to verbose messages.")]
        //    public bool Verbose { get; set; }
        //    [Option('x', "xml", Required = false, HelpText = "Set output format to XML.")]
        //    public bool Xml { get; set; }
        //    [Option('d', "diff", Required = false, HelpText = "Set text diff tool")]
        //    public string TextDiffTool { get; set; }
        //}

        //public enum OutputFormat
        //{
        //    json,
        //    xml
        //}

        static void Main(string[] args)
        {
            //bool isJson = true;
            //bool isXml = false;

            //Parser.Default.ParseArguments<Options>(args)
            //    .WithParsed<Options>(o =>
            //    {
            //        if (o.Verbose)
            //        {
            //            Console.WriteLine($"Verbose output enabled. Current Arguments: -v {o.Verbose}");
            //            Console.WriteLine("Quick Start Example! App is in Verbose mode!");
            //        }
            //        if (o.Xml)
            //        {
            //            isJson = false;
            //            isXml = true;
            //        }
            //        if (!string.IsNullOrEmpty(o.TextDiffTool))
            //        {
            //            Console.WriteLine($"You want to use {o.TextDiffTool}");
            //        }
            //    });


            //var report1 = new ReportDocument();

            //string report1Path = args[0];
            //string report2Path = args.Count() > 1 ? args[1] : "";

            

            //if (!File.Exists(report1Path))
            //{
            //    Console.WriteLine($"Can't seem to find {report1Path}");
            //    return;
            //}
            //if (Path.GetExtension(report1Path) != ".rpt")
            //{
            //    // TODO Pass filepath1 and filepath2 to textDiffTool if filepath2 is also not a report file
            //    Console.WriteLine($"{report1Path} isn't a report file");
            //    return;
            //}

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
                            && !File.Exists(args[1]))
                        {
                            SerializeToFile(args[0], args[1]);
                        }
                        else
                            Usage();
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
                            Usage();
                            return;
                        }

                        try
                        {
                            if (Path.GetExtension(rptFile1) == ".rpt")
                            {
                                file1IsRpt = true;
                                textFile1 = SerializeToFile(rptFile1);
                            }
                            else
                            {
                                textFile1 = rptFile1;
                            }

                            if (Path.GetExtension(rptFile2) == ".rpt")
                            {
                                file2IsRpt = true;
                                textFile2 = SerializeToFile(rptFile2);
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

                            Console.WriteLine($"Error: {ex.Message} {ex.InnerException.Message}");
                            var key = Console.ReadKey();
                        }
                        break;
                    }
                default:
                    {
                        Usage();
                        break;
                    }
            }

            //var filesToCompare = new List<DiffFile>();
            //try
            //{

            //foreach(var arg in args)
            //{
            //    if (arg == textDiffTool)
            //        continue;

            //    if (File.Exists(arg))
            //    {
            //        var diffFile = new DiffFile() { Path = arg };

            //        if (Path.GetExtension(arg) == ".rpt")
            //        {
            //            var rpt = new ReportDocument();
            //            diffFile.NewPath = Path.ChangeExtension(arg, "json");
            //            var file = File.CreateText(diffFile.NewPath);
            //            Console.WriteLine($"Loading {arg}");
            //            rpt.Load(arg);

            //            Console.WriteLine($"Saving {diffFile.NewPath}");
            //            diffFile.Serialized = Serialize(rpt, diffFile.NewPath);
            //            file.Write(diffFile.Serialized);
            //            file.Flush();
            //            file.Close();
            //            rpt.Close();
            //        }

            //        filesToCompare.Add(diffFile);
            //    }
            //}
            //}
            //catch (Exception ex)
            //{

            //    Console.WriteLine($"Error: {ex.Message} {ex.InnerException.Message}");
            //    var key = Console.ReadKey();
            //}
            //if (filesToCompare.Count() == 1)
            //{
            //    // output serialized file only
            //    Console.WriteLine("Blimey!");
            //    return;// filesToCompare.First();
            //}
            //else if (filesToCompare.Count() == 2)
            //{
            //    var diffParams = $"\"{filesToCompare[0].NewPath ?? filesToCompare[0].Path}\" \"{filesToCompare[1].NewPath ?? filesToCompare[1].Path}\"";
            //    Console.Write($"Starting \"{ textDiffTool}\" {diffParams}");

            //    var diffProc = Process.Start(
            //        textDiffTool, 
            //        diffParams
            //        );

            //    diffProc.WaitForExit();

            //    // Clean up (delete) temp files after text diff tool is finished with them
            //    foreach(var file in filesToCompare)
            //    {
            //        if (File.Exists(file.NewPath))
            //        {
            //            // TODO: Only if it was a temp file!
            //            File.Delete(file.NewPath);
            //        }
            //    }

            //    return;
            //}
            //else
            //{
            //    return;// "Eh?";
            //}
        }

        static void Usage()
        {
            Console.WriteLine("CRDiff PathToReport");
            Console.WriteLine("CRDiff PathToReport TargetFilename");
            Console.WriteLine("CRDiff PathToTextDiffApp PathToReport1 PathToReport2");
            Console.ReadKey();
        }

        static string SerializeToFile(string rptPath, string textPath = null)
        {
            var rpt = new ReportDocument();

            textPath = textPath ?? Path.ChangeExtension(rptPath, "json");
            var file = File.CreateText(textPath);
            Console.WriteLine($"Loading {rptPath}");
            rpt.Load(rptPath);

            Console.WriteLine($"Saving {textPath}");
            var serialized = Serialize(rpt, textPath);
            file.Write(serialized);
            file.Flush();
            file.Close();
            rpt.Close();
            return textPath;
        }
        static string Serialize(ReportDocument rpt, string filepath)//, OutputFormat fmt = OutputFormat.json)
        {
            var settings = new JsonSerializerSettings
            {
                ContractResolver = new CustomResolver(),
                Formatting = Formatting.Indented
            };

            string rptName = string.IsNullOrEmpty(rpt.Name) 
                ? Path.GetFileNameWithoutExtension(filepath) 
                : rpt.Name;

            var reportTables = rpt.ReportClientDocument.DatabaseController.Database.Tables;
            var commandSqlList = new StringBuilder();
            foreach (dynamic table in reportTables)
            {
                if (table.ClassName == "CrystalReports.CommandTable")
                {
                    commandSqlList.Append($"    {{\"ReportName\": \"{rptName}\", \n      \"Command\": [\n\"");
                    commandSqlList.Append(table.CommandText.Replace("\r\n", "\", \n\""));
                    commandSqlList.Append("\"]\n  }\n");
                }
            }

            var txt = new StringBuilder();
            txt.AppendLine("{\"Report\":\n  {\"CommandSQL\":");
            txt.Append(commandSqlList);
            txt.AppendLine("},");
            txt.AppendLine("{\"DatabaseTables\":");
            txt.AppendLine(JsonConvert.SerializeObject(rpt.ReportClientDocument.DatabaseController.Database.Tables, settings));
            txt.AppendLine("},");
            txt.AppendLine("{\"DataDefinition\":");
            txt.AppendLine(JsonConvert.SerializeObject(rpt.DataDefinition, settings));
            txt.AppendLine("},");
            txt.AppendLine("{\"ReportDefinition\":");
            txt.AppendLine(JsonConvert.SerializeObject(rpt.ReportDefinition, settings));
            txt.AppendLine("  }\n}");

            return txt.ToString();
        }

        static List<string> Parameters(ReportDocument rpt)
        {
            var paramtrs = new List<string>();
            var parameters = rpt.DataDefinition.ParameterFields;
            foreach (ParameterFieldDefinition param in parameters)
            {
                paramtrs.Add($"Parameter {param.ReportName}.{param.Name}");
            }
            return paramtrs;
        }
    }
    class CustomResolver : DefaultContractResolver
    {
        protected override JsonProperty CreateProperty(MemberInfo member, MemberSerialization memberSerialization)
        {
            JsonProperty property = base.CreateProperty(member, memberSerialization);

            property.ShouldSerialize = instance =>
            {
                try
                {
                    PropertyInfo prop = (PropertyInfo)member;
                    if (prop.CanRead)
                    {
                        prop.GetValue(instance, null);
                        return true;
                    }
                }
                catch (Exception)
                {
                }
                return false;
            };
            return property;
        }
    }

    //class DiffFile
    //{
    //    public string Path { get; set; }
    //    public string NewPath { get; set; }
    //    public string Serialized { get; set; }
    //    public string Destination { get; set; }
    //}
}
