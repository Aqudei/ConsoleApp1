using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json.Nodes;
using System.Threading.Tasks;

namespace ConsoleApp1.Exporter
{
    internal class JsonToExcelExporter
    {
        public JsonToExcelExporter(string input)
        {
            InputFile = input;
            processors = new List<ISheetProcessor>();
        }

        public string InputFile { get; }

        private List<ISheetProcessor> processors;

        public void RunExport(string destination)
        {

            var jsonText = File.ReadAllText(InputFile);
            var json = JsonNode.Parse(jsonText);
            var item = json.AsArray().FirstOrDefault();

            var projectNode = item["project"];
            var boreholesNode = item["boreholes"];

            processors.Add(new ProjectSheet(projectNode,boreholesNode));


            var workbook = new XLWorkbook();           


            foreach (var processor in processors)
            {
                var sheet = processor.Process();
                var worksheet = workbook.Worksheets.Add(processor.Name);
                worksheet.Cell(1,1).InsertTable(sheet);
            }
            workbook.SaveAs(destination);
        }
    }
}
