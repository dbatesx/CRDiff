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
     * 
     * Testing arguments:
     * "C:\Program Files (x86)\Compare It!\wincmp3.exe" "C:\dev\svn\Reports\Crystal Reports\trunk\KinderMorgan_CS_MT_Testing.rpt" "C:\dev\svn\Reports\Crystal Reports\trunk\KinderMorgan_CS_MT.rpt"
     * "C:\dev\svn\Reports\Crystal Reports\trunk\KinderMorgan_CS_MT_Testing.rpt" KM_CS_MT_Testing.json
     */
    class Program
    {

        static void Main(string[] args)
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

        }

        static void Usage(string[] args = null)
        {
            Console.WriteLine("CRDiff PathToReport");
            Console.WriteLine("CRDiff PathToReport TargetFilename");
            Console.WriteLine("CRDiff PathToTextDiffApp PathToReport1 PathToReport2");
            //Console.ReadKey();
        }

        static string SerializeToFile(string rptPath, string textPath = null, int reportOrder = 1)
        {
            var rpt = new ReportDocument();

            textPath = textPath ?? Path.ChangeExtension(rptPath, "json");
            var file = File.CreateText(textPath);
            Console.WriteLine($"Loading {rptPath}");
            rpt.Load(rptPath);

            Console.WriteLine($"Saving {textPath}");
            var serialized = Serialize(rpt, textPath, reportOrder);
            file.Write(serialized);
            file.Flush();
            file.Close();
            rpt.Close();
            return textPath;
        }
        static string Serialize(ReportDocument rpt, string filepath, int reportOrder = 1)//, OutputFormat fmt = OutputFormat.json)
        {
            var settings = new JsonSerializerSettings
            {
                ContractResolver = new CustomJsonResolver(),
                Formatting = Formatting.Indented
            };

            string rptName = string.IsNullOrEmpty(rpt.Name) 
                ? Path.GetFileNameWithoutExtension(filepath) 
                : rpt.Name;

            var reportTables = rpt.ReportClientDocument.DatabaseController.Database.Tables;
            var commandSqlList = new StringBuilder();
            var sw = new StringWriter();

            // TODO: Write property names and values with JsonWriter:
            //using (JsonWriter jw = new JsonTextWriter(sw))
            //{
            //    jw.WriteStartObject();

            //    jw.WritePropertyName("CrystalReports.CommandTable");
            //    jw.WritePropertyName("ReportName");
            //    jw.WriteValue(rptName);
            //    jw.WriteEndObject();

            //    commandSqlList.Append(jw);
            //}
            
            foreach (dynamic table in reportTables)
            {
                if (table.ClassName == "CrystalReports.CommandTable")
                {
                    commandSqlList.Append($"    {{\"ReportName\": \"{rptName}\", \n      \"Command\": [\n\"");
                    commandSqlList.Append(table.CommandText.Replace("\r\n", "\", \n\""));
                    commandSqlList.Append("\"]\n  }\n");
                }
            }
            // TODO: Need to show element suppression formulas (and other formulas for elements)
            // TODO: Need to show sub-textbox formatting (ie font changes, tab settings, etc)
            var txt = new StringBuilder();
            txt.AppendLine("{\"Report\":\n  {\"CommandSQL\":");
            txt.Append(commandSqlList);
            txt.AppendLine("},");
            txt.AppendLine("{\"DatabaseTables\":");
            txt.AppendLine(JsonConvert.SerializeObject(rpt.ReportClientDocument.DatabaseController.Database.Tables, settings));
            //if(reportOrder == 1)
            //{
            //    txt.AppendLine(JsonConvert.SerializeObject(rpt.Database, settings));
            //    txt.AppendLine(JsonConvert.SerializeObject(rpt.DataSourceConnections, settings));
            //    txt.AppendLine(JsonConvert.SerializeObject(rpt.Subreports, settings));

            //}
            txt.AppendLine("},");
            txt.AppendLine("{\"DataDefinition\":");
            txt.AppendLine(JsonConvert.SerializeObject(rpt.DataDefinition, settings));
            txt.AppendLine("},");
            txt.AppendLine("{\"ReportDefinition\":");
            txt.AppendLine(JsonConvert.SerializeObject(rpt.ReportDefinition, settings));
            // TODO: Add PageSetup info such as margins
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
}
