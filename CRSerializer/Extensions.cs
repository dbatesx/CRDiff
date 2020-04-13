using CrystalDecisions.CrystalReports.Engine;
//using CrystalDecisions;
using Newtonsoft.Json;
using System.Collections.Generic;

namespace CRSerializer
{
    internal static class Extensions
    {
        public static void WriteProperty(this JsonWriter jw, string name, object value)
        {
            jw.WritePropertyName(name);
            jw.WriteValue(value);
        }

        public static void WriteArray(this JsonWriter jw, string name, object[] value)
        {

        }

        public static void WriteObjectHierarchy(this JsonWriter jw, object obj)
        {
            var jsonSerializerSettings = new JsonSerializerSettings
            {
                ContractResolver = new CustomJsonResolver(),
                Formatting = Formatting.Indented
            };
            jw.WriteRawValue(JsonConvert.SerializeObject(obj, jsonSerializerSettings));
        }

        public static void WriteObjectHierarchy(this JsonWriter jw, string name, object obj)
        {
            jw.WritePropertyName(name);
            jw.WriteObjectHierarchy(obj);
        }

        public static void WriteDataSource(this JsonWriter jw, string keyName, CrystalDecisions.ReportAppServer.DataDefModel.Tables tables)
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

        public static void WriteParameters(this JsonWriter jw, string keyName, ReportDocument rpt)
        {
            //var parameterList = new List<string>();
            var parameterFields = rpt.DataDefinition.ParameterFields;
            jw.WritePropertyName(keyName);
            jw.WriteStartArray();
            foreach (ParameterFieldDefinition parameterField in parameterFields)
            {
                jw.WriteStartObject();
                jw.WriteProperty("Name", parameterField.Name);
                jw.WriteProperty("FormulaName", parameterField.FormulaName);
                jw.WriteProperty("ValueType", parameterField.ValueType.ToString());
                jw.WriteProperty("EnableNullValue", parameterField.EnableNullValue.ToString());
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

        public static void WritePrintOptions(this JsonWriter jw, PrintOptions printOptions)
        {
            jw.WritePropertyName("PrintOptions");
            jw.WriteStartObject();
            jw.WriteProperty("NoPrinter", printOptions.NoPrinter);
            jw.WriteProperty("PrinterName", printOptions.PrinterName);
            jw.WriteProperty("SavedPrinterName", printOptions.SavedPrinterName);
            jw.WriteProperty("PrinterDuplex", printOptions.PrinterDuplex.ToString());
            jw.WriteProperty("PaperOrientation", printOptions.PaperOrientation.ToString());
            jw.WriteProperty("PaperSize", printOptions.PaperSize.ToString());
            jw.WriteProperty("PaperSource", printOptions.PaperSource.ToString());
            jw.WriteProperty("DissociatePageSizeAndPrinterPaperSize", printOptions.DissociatePageSizeAndPrinterPaperSize);
            jw.WriteProperty("PageContentHeight", $"{printOptions.PageContentHeight} twips ({printOptions.PageContentHeight / 1440.0:N3} inches)");
            jw.WriteProperty("PageContentWidth", $"{printOptions.PageContentWidth} twips ({printOptions.PageContentWidth / 1440.0:N3} inches)");

            jw.WritePropertyName("PageMargins");
            jw.WriteStartObject();
            jw.WriteProperty("topMargin", $"{printOptions.PageMargins.topMargin} twips ({printOptions.PageMargins.topMargin / 1440.0:N3} inches)");
            jw.WriteProperty("leftMargin", $"{printOptions.PageMargins.leftMargin} twips ({printOptions.PageMargins.leftMargin / 1440.0:N3} inches)");
            jw.WriteProperty("rightMargin", $"{printOptions.PageMargins.rightMargin} twips ({printOptions.PageMargins.rightMargin / 1440.0:N3} inches)");
            jw.WriteProperty("bottomMargin", $"{printOptions.PageMargins.bottomMargin} twips ({printOptions.PageMargins.bottomMargin / 1440.0:N3} inches)");

            jw.WriteEndObject();
        }
    }
}