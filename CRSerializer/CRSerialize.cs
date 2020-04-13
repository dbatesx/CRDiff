using CrystalDecisions.CrystalReports.Engine;
using Newtonsoft.Json;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace CRSerializer
{
    public class CRSerialize
    {
        public string Serialize(string rptPath, string filepath = null, int reportOrder = 1)//, OutputFormat fmt = OutputFormat.json)
        {
            var rpt = new ReportDocument();
            rpt.Load(rptPath);
            var rptClient = rpt.ReportClientDocument;

            var jsonSerializerSettings = new JsonSerializerSettings
            {
                ContractResolver = new CustomJsonResolver(),
                Formatting = Formatting.Indented
            };

            string rptName = string.IsNullOrEmpty(rpt.Name)
                ? Path.GetFileNameWithoutExtension(rptPath)
                : rpt.Name;

            var sb = new StringBuilder();
            var sw = new StringWriter(sb);

            using (JsonWriter jw = new JsonTextWriter(sw))
            {
                jw.Formatting = Formatting.Indented;

                jw.WriteStartObject();

                jw.WriteObjectHierarchy("Entire Report Object", rpt);

                jw.WriteProperty("SerializeVersion", CRSerializeProductVersion());
                
                jw.WriteProperty("ReportName", rptName);

                jw.WriteParameters("Parameters", rpt);

                jw.WriteObjectHierarchy("SummaryInfo", rpt.SummaryInfo);

                // experimental properties:
                jw.WriteObjectHierarchy("DataSourceConnections", rpt.DataSourceConnections);

                jw.WriteObjectHierarchy("ParameterFields", rpt.ParameterFields);

                jw.WriteObjectHierarchy("ReportOptions", rptClient.ReportOptions);

                jw.WriteDataSource("DataSource", rptClient.Database.Tables);

                jw.WriteObjectHierarchy("DatabaseTables", rpt.Database.Tables);

                jw.WriteObjectHierarchy("DataDefinition", rpt.DataDefinition);

                jw.WritePrintOptions(rpt.PrintOptions);

                jw.WriteObjectHierarchy("CustomFunctions", rptClient.CustomFunctionController.GetCustomFunctions());
                
                jw.WriteObjectHierarchy("ReportDefinition", rpt.ReportDefinition);

                jw.WritePropertyName("Subreports");
                var subReports = rpt.Subreports;
                if (subReports.Count > 0)
                {
                    jw.WriteStartArray();
                    foreach (ReportDocument subReport in subReports)
                    {
                        var subReportClient = rptClient.SubreportController.GetSubreport(subReport.Name);

                        jw.WriteStartObject();
                        jw.WriteProperty("SubreportName", subReport.Name);

                        jw.WriteParameters("Parameters", subReport);

                        jw.WriteDataSource("DataSource", subReportClient.DataDefController.Database.Tables);

                        jw.WriteObjectHierarchy("DatabaseTables", subReport.Database.Tables);

                        jw.WriteObjectHierarchy("DataDefinition", subReport.DataDefinition);

                        jw.WriteObjectHierarchy("CustomFunctions", rptClient.CustomFunctionController.GetCustomFunctions());

                        jw.WriteObjectHierarchy("ReportDefinition", subReport.ReportDefinition);

                        jw.WriteEndObject();
                    }
                }

                jw.WriteEndObject(); // final end
                rpt.Close();

                return sb.ToString();
            }
        }
        public static string CRSerializeProductVersion()
        {
            return FileVersionInfo
                .GetVersionInfo(
                    Assembly
                    .GetExecutingAssembly()
                    .Location
                    )
                .ProductVersion;
        }

        public string CRVersion()
        {
            // ref: https://apps.support.sap.com/sap/support/knowledge/public/en/2003551
            var versions = new StringBuilder();
            var rpt = new ReportDocument(); // pull in Crystal library

            foreach (var crAsm in AppDomain.CurrentDomain.GetAssemblies()
                .Where(a => a.FullName.StartsWith("CrystalDecisions.CrystalReports.Engine")))
            {
                var fileVersionInfo = System.Diagnostics.FileVersionInfo.GetVersionInfo(crAsm.Location);
                versions.Append(fileVersionInfo.FileVersion);
                //versions.Append("\n");
                //versions.Append(fileVersionInfo.FileName);
            }
            if (versions.Length == 0)
            {
                versions.Append("Not Found");
            }
            return versions.ToString();
        }


    }
}