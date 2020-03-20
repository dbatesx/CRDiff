using CrystalDecisions.CrystalReports.Engine;
using Newtonsoft.Json;
using System;
using System.IO;
using System.Linq;
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
                ? Path.GetFileNameWithoutExtension(filepath)
                : rpt.Name;

            var sb = new StringBuilder();
            var sw = new StringWriter(sb);

            using (JsonWriter jw = new JsonTextWriter(sw))
            {
                jw.Formatting = Formatting.Indented;

                jw.WriteStartObject();

                jw.WritePropertyName("ReportName");
                jw.WriteValue(rptName);

                jw.WritePropertyName("DataSource");
                jw.WriteDataSource(rptClient.Database.Tables);

                jw.WritePropertyName("DatabaseTables");
                jw.WriteObjectHierarchy(rptClient.Database.Tables);

                jw.WritePropertyName("DataDefinition");
                jw.WriteObjectHierarchy(rpt.DataDefinition);

                jw.WritePropertyName("ReportDefinition");
                jw.WriteObjectHierarchy(rpt.ReportDefinition);
                //jw.WriteRawValue(JsonConvert.SerializeObject(rpt.ReportDefinition.ReportObjects, jsonSerializerSettings));

                jw.WritePropertyName("Subreports");
                var subReports = rpt.Subreports;
                if (subReports.Count > 0)
                {
                    jw.WriteStartArray();
                    foreach (ReportDocument subReport in subReports)
                    {
                        var subReportClient = rptClient.SubreportController.GetSubreport(subReport.Name);

                        jw.WriteStartObject();
                        jw.WritePropertyName("SubreportName");
                        jw.WriteValue(subReport.Name);

                        jw.WritePropertyName("DataSource");
                        jw.WriteDataSource(subReportClient.DataDefController.Database.Tables);

                        jw.WritePropertyName("DatabaseTables");
                        jw.WriteObjectHierarchy(subReport.Database.Tables);

                        jw.WritePropertyName("DataDefinition");
                        jw.WriteObjectHierarchy(subReport.DataDefinition);

                        jw.WritePropertyName("ReportDefinition");
                        jw.WriteObjectHierarchy(subReport.ReportDefinition);

                        jw.WriteEndObject();
                    }
                }

                jw.WriteEndObject(); // final end
                rpt.Close();

                return sb.ToString();
            }
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