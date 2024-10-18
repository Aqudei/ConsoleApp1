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
            if (json == null)
            {
                Console.WriteLine("Cannot parse input file to JSON");
                return;
            }

            var item = json.AsArray().FirstOrDefault();

            var projectNode = item?["project"];
            var boreholesNode = item?["boreholes"];

            processors.Add(new SheetProject(projectNode, boreholesNode));
            processors.Add(new SheetTestHole(projectNode, boreholesNode));
            processors.Add(new SheetBackFill(projectNode, boreholesNode));
            processors.Add(new SheetPiezometers(projectNode, boreholesNode));
            processors.Add(new SheetFieldTests(projectNode, boreholesNode));
            processors.Add(new SheetComments(projectNode, boreholesNode));
            processors.Add(new SheetDrillRuns(projectNode, boreholesNode));
            processors.Add(new SheetDiscontinuities(projectNode, boreholesNode));
            processors.Add(new SheetDrillingDetails(projectNode, boreholesNode));

            var workbook = new XLWorkbook();


            foreach (var processor in processors)
            {
                var sheet = processor.Process();
                var worksheet = workbook.Worksheets.Add(processor.Name);
                if (sheet == null)
                    continue;

                worksheet.Cell(1, 1).InsertTable(sheet);
            }
            workbook.SaveAs(destination);
        }
    }
}
