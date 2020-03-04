using CrystalDecisions.CrystalReports.Engine;
using Newtonsoft.Json;
using System;
using System.IO;
using System.Text;
using System.Linq;

namespace CRSerializer
{
    public class CRSerialize
    {
        public string Serialize(ReportDocument rpt, string filepath, int reportOrder = 1)//, OutputFormat fmt = OutputFormat.json)
        {
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
                GetDataSource(jw, rptClient.Database.Tables);
                jw.WritePropertyName("DatabaseTables");
                jw.WriteRawValue(JsonConvert.SerializeObject(rptClient.Database.Tables, jsonSerializerSettings));
                jw.WritePropertyName("DataDefinition");
                jw.WriteRawValue(JsonConvert.SerializeObject(rpt.DataDefinition, jsonSerializerSettings));
                jw.WritePropertyName("ReportDefinition");
                jw.WriteRawValue(JsonConvert.SerializeObject(rpt.ReportDefinition, jsonSerializerSettings));
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
                        GetDataSource(jw, subReportClient.DataDefController.Database.Tables);
                        jw.WritePropertyName("DatabaseTables");
                        //WriteObject(jw, subReport.Database.Tables);
                        jw.WriteRawValue(JsonConvert.SerializeObject(subReport.Database.Tables, jsonSerializerSettings));
                        jw.WritePropertyName("DataDefinition");
                        jw.WriteRawValue(JsonConvert.SerializeObject(subReport.DataDefinition, jsonSerializerSettings));
                        jw.WritePropertyName("ReportDefinition");
                        jw.WriteRawValue(JsonConvert.SerializeObject(subReport.ReportDefinition, jsonSerializerSettings));
                        jw.WriteEndObject();
                    }
                }

                jw.WriteEndObject(); // final end

                return sb.ToString();
            }
        }

        private void WriteObject(JsonWriter jw, object obj)
        {
            var jsonSerializerSettings = new JsonSerializerSettings
            {
                ContractResolver = new CustomJsonResolver(),
                Formatting = Formatting.Indented
            };
            jw.WriteRawValue(JsonConvert.SerializeObject(obj, jsonSerializerSettings));

        }
        private void GetDataSource(JsonWriter jw, CrystalDecisions.ReportAppServer.DataDefModel.Tables tables)
        {
            //Doesn't deal with a mix of command statements and tables (but that scenario is rare)
            var isFirstTable = true;

            jw.WriteStartObject();
            foreach (dynamic table in tables)
            {
                if (table.ClassName == "CrystalReports.CommandTable")
                {
                    jw.WritePropertyName("Command");
                    MultiLinesToArray(jw, table.CommandText);
                }
                else
                {
                    if (isFirstTable)
                    {
                        jw.WritePropertyName("Tables");
                        jw.WriteStartArray();
                        isFirstTable = false;
                    }
                    jw.WriteValue(table.Name);
                }
            }
            if (!isFirstTable)
            {
                jw.WriteEndArray();
            }
            jw.WriteEndObject();
        }

        private void MultiLinesToArray(JsonWriter w, string str)
        {
            string[] vs = str.Split(new[] { "\r\n", "\n" }, System.StringSplitOptions.None);
            w.WriteStartArray();
            foreach (var s in vs)
            {
                w.WriteValue(s);
            }
            w.WriteEnd();
        }
    }
}