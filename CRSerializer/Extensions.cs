using CrystalDecisions.CrystalReports.Engine;
//using CrystalDecisions;
using Newtonsoft.Json;
using System.Collections.Generic;

namespace CRSerializer
{
    internal static class Extensions
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

        public static void WriteParameters(this JsonWriter jw, ReportDocument rpt)
        {
            //var parameterList = new List<string>();
            var parameterFields = rpt.DataDefinition.ParameterFields;
            //jw.WritePropertyName("Parameters");
            jw.WriteStartArray();
            foreach (ParameterFieldDefinition parameterField in parameterFields)
            {
                jw.WriteStartObject();
                jw.WritePropertyName("Name");
                jw.WriteValue(parameterField.Name);
                jw.WritePropertyName("FormulaName");
                jw.WriteValue(parameterField.FormulaName);
                jw.WritePropertyName("ValueType");
                jw.WriteValue(parameterField.ValueType.ToString());
                jw.WritePropertyName("EnableNullValue");
                jw.WriteValue(parameterField.EnableNullValue.ToString());
                jw.WriteEndObject();
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