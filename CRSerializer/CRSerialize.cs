using CrystalDecisions.CrystalReports.Engine;
using Newtonsoft.Json;
using System;
using System.IO;
using System.Text;

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

            // TODO: Write property names and values with JsonWriter:
            using (JsonWriter jw = new JsonTextWriter(sw))
            {
                jw.Formatting = Formatting.Indented;
                //jw.

                jw.WriteStartObject();

                //jw.WritePropertyName("CrystalReports.CommandTable");
                jw.WritePropertyName("ReportPropertyName");
                jw.WriteValue(rptName);
                jw.WriteEndObject();

                //commandSqlList.Append(jw.);
                GetTables(jw, rptClient.Database.Tables);

                //var subReports = rpt.Subreports;
                //foreach (ReportDocument subReport in subReports)
                //{
                //    CommandTable(jw, subReport.Database.Tables);
                //    //subReport.Database.Tables;

                //}

                var subReportNames = rptClient.SubreportController.GetSubreportNames();
                foreach (string subreportName in subReportNames)
                {
                    var subReport = rptClient.SubreportController.GetSubreport(subreportName);
                    GetTables(jw, subReport.DataDefController.Database.Tables);
                }

                //foreach (dynamic table in reportTables)
                //{
                //    if (table.ClassName == "CrystalReports.CommandTable")
                //    {
                //        commandSqlList.Append($"    {{\"ReportName\": \"{rptName}\", \n      \"Command\": [\n\"");
                //        commandSqlList.Append(table.CommandText.Replace("\r\n", "\", \n\""));
                //        commandSqlList.Append("\"]\n  }\n");
                //    }
                //}

                var txt = new StringBuilder();
                txt.AppendLine("{\"Report\":\n  {\"CommandSQL\":");
                txt.Append(sb);
                txt.AppendLine("},");
                txt.AppendLine("{\"DatabaseTables\":");
                txt.AppendLine(JsonConvert.SerializeObject(rpt.ReportClientDocument.DatabaseController.Database.Tables, jsonSerializerSettings));
                //if(reportOrder == 1)
                //{
                //    txt.AppendLine(JsonConvert.SerializeObject(rpt.Database, settings));
                //    txt.AppendLine(JsonConvert.SerializeObject(rpt.DataSourceConnections, settings));
                //    txt.AppendLine(JsonConvert.SerializeObject(rpt.Subreports, settings));

                //}
                txt.AppendLine("},");
                txt.AppendLine("{\"DataDefinition\":");
                txt.AppendLine(JsonConvert.SerializeObject(rpt.DataDefinition, jsonSerializerSettings));
                txt.AppendLine("},");
                txt.AppendLine("{\"ReportDefinition\":");
                txt.AppendLine(JsonConvert.SerializeObject(rpt.ReportDefinition.Areas, jsonSerializerSettings));
                txt.AppendLine(JsonConvert.SerializeObject(rpt.ReportDefinition.ReportObjects, jsonSerializerSettings));
                // TODO: Add PageSetup info such as margins
                txt.AppendLine("  }\n}");

                return txt.ToString();
            }
        }

        private void CommandTable(JsonWriter w, CrystalDecisions.CrystalReports.Engine.Tables tables)
        {
            foreach (dynamic table in tables)
            {
                if (table.ClassName == "CrystalReports.CommandTable")
                {
                    w.WriteStartObject();
                    w.WritePropertyName("Command");
                    MultiLinesToArray(w, table.CommandText);
                    w.WriteEndObject();
                }
            }
        }

        private void GetTables(JsonWriter w, CrystalDecisions.ReportAppServer.DataDefModel.Tables tables)
        {
            foreach (dynamic table in tables)
            {
                w.WriteStartObject();
                if (table.ClassName == "CrystalReports.CommandTable")
                {
                    w.WritePropertyName("Command");
                    MultiLinesToArray(w, table.CommandText);
                }
                else
                {
                    w.WritePropertyName("Table");
                    w.WriteValue(table.Name);
                }
                w.WriteEndObject();
            }
        }

        private void MultiLinesToArray(JsonWriter w, string str)
        {
            //string[] lineEnd = { "\r\n", "\n" };
            string[] vs = str.Split(new[] { "\r\n", "\n" }, System.StringSplitOptions.None);
            w.WriteStartArray();
            foreach (var s in vs)
            {
                w.WriteValue(s);
            }
            w.WriteEnd();
        }

        private string getSQL(string rptName)
        {
            CrystalDecisions.CrystalReports.Engine.ReportDocument rpt = new CrystalDecisions.CrystalReports.Engine.ReportDocument();
            CrystalDecisions.ReportAppServer.ClientDoc.ISCDReportClientDocument rptClient;
            CrystalDecisions.ReportAppServer.Controllers.DataDefController dataDefController;
            CrystalDecisions.ReportAppServer.DataDefModel.Database boDatabase;
            CrystalDecisions.ReportAppServer.DataDefModel.CommandTable cmdTable;

            // Load the report using the CR .NET SDK and get a handle on the ReportClientDocument
            //boReportDocument.Load(iObject, eSession);
            rpt.Load(rptName);
            rptClient = rpt.ReportClientDocument;

            // Use the DataDefController to access the database and the command table.
            // Then display the current command table SQL in the textbox.
            dataDefController = rptClient.DataDefController;
            boDatabase = dataDefController.Database;

            string sql = "";

            for (int i = 0; i < boDatabase.Tables.Count; i++)
            {
                CrystalDecisions.ReportAppServer.DataDefModel.ISCRTable tableObject = boDatabase.Tables[i];

                if (tableObject.ClassName == "CrystalReports.Table")
                {
                    sql = sql + "Table " + i + ": " + tableObject.Name;
                }
                else
                {
                    cmdTable = (CrystalDecisions.ReportAppServer.DataDefModel.CommandTable)boDatabase.Tables[i];
                    sql = sql + "Query " + i + ": " + cmdTable.CommandText;
                }
                sql += Environment.NewLine;
            }

            foreach (string subName in rptClient.SubreportController.GetSubreportNames())
            {
                CrystalDecisions.ReportAppServer.Controllers.SubreportClientDocument subRCD = rptClient.SubreportController.GetSubreport(subName);

                for (int i = 0; i < boDatabase.Tables.Count; i++)
                {
                    CrystalDecisions.ReportAppServer.DataDefModel.ISCRTable tableObject = boDatabase.Tables[i];

                    if (tableObject.ClassName == "CrystalReports.Table")
                    {
                        sql = sql + "Table " + i + ": " + tableObject.Name;
                    }
                    else
                    {
                        cmdTable = (CrystalDecisions.ReportAppServer.DataDefModel.CommandTable)subRCD.DatabaseController.Database.Tables[i];
                        sql = sql + "Subreport " + subName + " - Query " + i + ": " + cmdTable.CommandText;
                    }
                    sql += Environment.NewLine;
                }
            }

            // Clean up
            return sql;
        }
    }
}