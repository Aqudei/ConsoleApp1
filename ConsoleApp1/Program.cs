


// See https://aka.ms/new-console-template for more information
using ClosedXML.Excel;
using ConsoleApp1;
using ConsoleApp1.Exporter;
using System.Data;
using System.Text.Json.Nodes;
using System.Text.Json.Serialization;

#region main

var inputFile = args[0];
var destination = Path.ChangeExtension(inputFile, "xlsx");
Console.WriteLine($"Input: {inputFile}");

var exporter = new JsonToExcelExporter(inputFile);
exporter.RunExport(destination);


#endregion