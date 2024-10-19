using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Diagnostics;
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
            //processors.Add(new SheetFieldTests(projectNode, boreholesNode));
            processors.Add(new SheetComments(projectNode, boreholesNode));
            processors.Add(new SheetDrillRuns(projectNode, boreholesNode));
            processors.Add(new SheetDiscontinuities(projectNode, boreholesNode));
            processors.Add(new SheetDrillingDetails(projectNode, boreholesNode));
            processors.Add(new SheetStratigraphy(projectNode, boreholesNode));
            processors.Add(new SheetSamples(projectNode, boreholesNode));

            var workbook = new XLWorkbook();


            foreach (var processor in processors)
            {
                var sheet = processor.Process();
                var worksheet = workbook.Worksheets.Add(processor.Name);
                if (sheet == null)
                    continue;

                worksheet.Cell(1, 1).InsertTable(sheet);
                worksheet.Columns().AdjustToContents();
            }


            var fieldTestProc = new SheetFieldTests(projectNode, boreholesNode);
            var ftTables = fieldTestProc.Process().ToArray();
            var worksheetFieldTests = workbook.Worksheets.Add("Field Tests");
            var currentColumn = 1;
            for (int i = 0; i < ftTables.Count(); i++)
            {
                var table = ftTables[i];
                worksheetFieldTests.Cell(1, currentColumn).InsertTable(table);

                //worksheetFieldTests.Cell(1, currentColumn).Value = table.TableName;
                //var range = $"{(char)(64 + currentColumn)}1:{(char)(64 + currentColumn+1)}1";
                //worksheetFieldTests.Range(range).Merge();
                
                worksheetFieldTests.Columns().AdjustToContents();
                currentColumn += table.Columns.Count + 1;
            }

            workbook.SaveAs(destination);
        }
    }
}
