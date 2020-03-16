using CrystalDecisions.CrystalReports.Engine;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CRSerializer
{
    static class Extensions
    {
        public static void WriteObjectHierarchy(this JsonWriter jw, object obj)
        {
            var jsonSerializerSettings = new JsonSerializerSettings
            {
                ContractResolver = new CustomJsonResolver(),
                Formatting = Formatting.Indented
            };
            jw.WriteRawValue(JsonConvert.SerializeObject(obj, jsonSerializerSettings));
        }

        public static void WriteDataSource(this JsonWriter jw, CrystalDecisions.ReportAppServer.DataDefModel.Tables tables)
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

        public static void MultiLinesToArray(this JsonWriter jw, string str)
        {
            string[] vs = str.Split(new[] { "\r\n", "\n" }, System.StringSplitOptions.None);
            jw.WriteStartArray();
            foreach (var s in vs)
            {
                jw.WriteValue(s);
            }
            jw.WriteEnd();
        }

        private static List<string> Parameters(ReportDocument rpt)
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
