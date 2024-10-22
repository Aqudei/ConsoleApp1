using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Text.Json.Nodes;
using System.Threading.Tasks;

namespace ConsoleApp1.Exporter
{
    public class SheetPiezometers : SheetProcessorBase, ISheetProcessor
    {
        public SheetPiezometers(JsonNode? projectNode, JsonNode? boreholesNode) : base(projectNode, boreholesNode)
        { }

        protected override IEnumerable<string> GetIgnoredProperties()
        {
            return new[] { "endCapTypeNumber" };
        }

        public string Name => "Piezometers";

        public IEnumerable<DataTable> Process()
        {
            var backFills = ProcessBoreHolesItems("piezometerData", "backfill");

            var piezometer = ProcessBoreHolesItems("piezometerData", "piezometer");

            var filtered = piezometer.Where(p =>
            {
                if (p.TryGetValue("testHole", out var piezometerTestHole))
                {
                    if (backFills.Any(b =>
                    {
                        if (b.TryGetValue("testHole", out var backFillTestHole))
                        {
                            var exist = backFillTestHole?.ToString() == piezometerTestHole?.ToString();
                            return exist;
                        }
                        else
                        {
                            return false;
                        }
                    })) return true;

                    return false;
                }
                return false;
            });

            yield return ToDataTable(filtered);
        }
    }

    public class SheetBackFill : SheetProcessorBase, ISheetProcessor
    {
        public SheetBackFill(JsonNode? projectNode, JsonNode? boreholesNode) : base(projectNode, boreholesNode)
        {
        }

        public string Name => "BackFill";

        public IEnumerable<DataTable> Process()
        {
            yield return ToDataTable(ProcessBoreHolesItems("piezometerData", "backfill"));
        }
    }
    public class SheetDrillingDetails : SheetProcessorBase, ISheetProcessor
    {
        public SheetDrillingDetails(JsonNode? projectNode, JsonNode? boreholesNode) : base(projectNode, boreholesNode)
        {
        }

        public string Name => "Drilling Details";

        public override void FormatSheet(IXLWorksheet worksheet)
        {
            ResizeAndWrapColumn(worksheet, "Notes", 50);
        }

        public IEnumerable<DataTable> Process()
        {
            yield return ToDataTable(ProcessBoreHolesItems("boringMethods"));
        }
    }

    public class SheetStratigraphy : SheetProcessorBase, ISheetProcessor
    {
        public SheetStratigraphy(JsonNode? projectNode, JsonNode? boreholesNode) : base(projectNode, boreholesNode)
        { }

        public string Name => "Stratigraphy";

        public override void FormatSheet(IXLWorksheet worksheet)
        {
            ResizeAndWrapColumn(worksheet, "Description", 80);
        }

        protected override IEnumerable<string> GetIgnoredProperties()
        {
            return new[] { "dataEntryMode", "", "classificationDetailsSoil", "classificationDetailsRock" };
        }

        protected override Dictionary<string, string>? GetColumnsMapping()
        {
            return new Dictionary<string, string> { { "Unit", "Geol. Unit" }, { "Category", "Geol. Category" } };
        }

        public IEnumerable<DataTable> Process()
        {
            yield return ToDataTable(ProcessBoreHolesItems("stratigraphy", include: new string[] { "geologicUnit", "classificationSystem" }));
        }
    }
    public class SheetDiscontinuities : SheetProcessorBase, ISheetProcessor
    {
        public SheetDiscontinuities(JsonNode? projectNode, JsonNode? boreholesNode) : base(projectNode, boreholesNode)
        {
        }

        public string Name => "Discontinuities";

        public override void FormatSheet(IXLWorksheet worksheet)
        {
            ResizeAndWrapColumn(worksheet, "Infill Type", 32);
            ResizeAndWrapColumn(worksheet, "Infill Character", 32);
            ResizeAndWrapColumn(worksheet, "Infill Thickness", 32);
        }

        public IEnumerable<DataTable> Process()
        {
            yield return ToDataTable(ProcessBoreHolesItems("discontinuities"));
        }
    }

    public class SheetDrillRuns : SheetProcessorBase, ISheetProcessor
    {
        public SheetDrillRuns(JsonNode? projectNode, JsonNode? boreholesNode) : base(projectNode, boreholesNode)
        { }

        public string Name => "Drill Runs";

        protected override Dictionary<string, string>? GetColumnsMapping()
        {
            return new Dictionary<string, string> { { "Total Length", "Total Length of >4 inch Segments" } };
        }

        public IEnumerable<DataTable> Process()
        {
            var unit = _unitSystem != null && _unitSystem == "Imperial" ? "in" : "cm";


            var table = ToDataTable(ProcessBoreHolesItems("drillRuns"));

            foreach (DataColumn col in table.Columns)
            {
                if (col.ColumnName.Contains("Rqd Rmu") ||
                    col.ColumnName.Contains("Rqd Core Length") ||
                    col.ColumnName.Contains("Rmr") ||
                    col.ColumnName.Contains("Tcr") ||
                    col.ColumnName.Contains("Tmr") ||
                    col.ColumnName.Contains("Scr"))
                {
                    col.ColumnName = $"{col.ColumnName} (%)";
                }


                if (col.ColumnName.Contains("Total Core Recovered") ||
                   col.ColumnName.Contains("Total Length "))
                {
                    col.ColumnName = $"{col.ColumnName} ({unit})";
                }
            }

            yield return table;
        }
    }

    public class SheetSamples : SheetProcessorBase, ISheetProcessor
    {
        public SheetSamples(JsonNode? projectNode, JsonNode? boreholesNode) : base(projectNode, boreholesNode)
        {
        }

        public string Name => "Samples";

        protected override Dictionary<string, string>? GetColumnsMapping()
        {
            return new Dictionary<string, string> { { "Number", "Sample No." } };
        }


        protected override IEnumerable<string> GetIgnoredProperties() => new string[] { "typeIconUrl" };

        public IEnumerable<DataTable> Process()
        {
            var items = new List<Dictionary<string, object>>();

            foreach (var boreHole in _boreholesNode?.AsArray() ?? new JsonArray())
            {
                var samples = boreHole?["samples"];
                var name = boreHole?["name"];

                if (samples == null)
                    continue;

                foreach (var sample in samples.AsArray())
                {
                    var data = ExtractProperties(sample);

                    data["testHole"] = name;

                    var labTests = sample?["labTests"];

                    foreach (var labTest in labTests.AsObject())
                    {
                        foreach (var item in labTest.Value.AsArray())
                        {
                            var props = ExtractProperties(item);
                            foreach (var prop in props)
                            {
                                if (prop.Value is JsonValue || prop.Value is null)
                                    data[prop.Key] = prop.Value;
                            }

                            break;
                        }
                    }

                    items.Add(data);
                }
            }

            yield return ToDataTable(items);
        }
    }
    public class SheetTestHole : SheetProcessorBase, ISheetProcessor
    {
        public string Name => "Test Holes";

        public SheetTestHole(JsonNode? projectNode, JsonNode? boreholesNode) : base(projectNode, boreholesNode)
        {
        }


        protected override Dictionary<string, int> GetColumnsOrdering()
        {
            return new Dictionary<string, int> {
                { "Name", 0 },
                { "Test Hole Type", 1 },
                { $"Depth ({_unit})", 2 },
                { $"Groundwater Depth ({_unit})", 3 },
                { $"Groundwater Elev ({_crsMeasurementUnit})", 4 }
            };
        }

        public IEnumerable<DataTable> Process()
        {
            var items = new List<Dictionary<string, object>>();

            foreach (var boreHole in _boreholesNode.AsArray())
            {
                var data = new Dictionary<string, object>();
                foreach (var property in boreHole.AsObject())
                {
                    if (ShouldSkipKey(property.Key))
                        continue;

                    if (property.Key == "drillingGroundwaterLevels")
                    {
                        data["groundwaterDepth"] = property.Value["groundwaterDepth"];
                        data["groundwaterElev"] = property.Value["groundwaterElev"];
                    }

                    else if (property.Key == "progressStatus")
                    {
                        data["progressStatus"] = property.Value["name"];
                    }

                    else if (property.Key == "completionNotes")
                    {
                        foreach (var completionNotesNode in property.Value.AsObject())
                        {
                            data[completionNotesNode.Key] = completionNotesNode.Value;
                            data[completionNotesNode.Key] = completionNotesNode.Value;
                        }
                    }

                    else if (property.Key == "sptHammer")
                    {
                        foreach (var sptHammerNode in property.Value.AsObject())
                        {
                            data[$"SPT{sptHammerNode.Key}"] = sptHammerNode.Value;
                            data[$"SPT{sptHammerNode.Key}"] = sptHammerNode.Value;
                        }
                    }
                    else
                    {
                        if (property.Value is JsonValue jsonValue || property.Value is null)
                        {
                            if (property.Key == "northing" || property.Key == "easting")
                            {
                                var decimalValue = property.Value?.GetValue<decimal>();

                                if (decimalValue != null && decimalValue.HasValue)
                                {
                                    data[property.Key] = decimalValue.Value.ToString("N0");
                                }
                                else
                                {
                                    data[property.Key] = string.Empty;
                                }

                                continue;
                            }

                            data[property.Key] = property.Value;
                        }
                    }
                }

                items.Add(data);
            }

            yield return ToDataTable(items);
        }


    }



    public class SheetComments : SheetProcessorBase, ISheetProcessor
    {
        public SheetComments(JsonNode? projectNode, JsonNode? boreholesNode) : base(projectNode, boreholesNode)
        {
        }

        public override void FormatSheet(IXLWorksheet worksheet)
        {
            ResizeAndWrapColumn(worksheet, "Description", 50);
        }

        public string Name => "Comments";

        public IEnumerable<DataTable> Process()
        {
            yield return ToDataTable(ProcessBoreHolesItems("comments"));
        }
    }

    public class SheetFieldTests : SheetProcessorBase, ISheetProcessor
    {
        public SheetFieldTests(JsonNode? projectNode, JsonNode? boreholesNode) : base(projectNode, boreholesNode)
        {
        }

        protected override Dictionary<string, int> GetColumnsOrdering()
        {
            return new Dictionary<string, int> { { "Test Hole", 0 }, { "Test Title", 1 }, { "Depth (ft)", 2 }, { "Value", 3 } };
        }

        protected override IEnumerable<string> GetIgnoredProperties()
        {
            return new string[] { "action", "summary" };
        }

        public string Name => "Field Tests";

        public IEnumerable<DataTable> Process()
        {
            var existingTables = new List<DataTable>();

            foreach (var borehole in _boreholesNode.AsArray())
            {
                var testHole = borehole["name"];

                var fieldTests = borehole["fieldTests"];
                foreach (var field in fieldTests.AsArray())
                {
                    var items = new List<Dictionary<string, object>>();
                    var columns = field?["columns"]?.ToString().Split(',');
                    var fieldProps = ExtractProperties(field);
                    fieldProps.Remove("depth", out var _);
                    fieldProps.Remove("columns", out var _);

                    foreach (var val in field["values"].AsArray())
                    {
                        var valueProps = ExtractProperties(val);
                        var data = new Dictionary<string, object>();
                        data["testHole"] = testHole;

                        foreach (var prop in fieldProps)
                        {
                            if (ShouldSkipKey(prop.Key))
                                continue;


                            data[prop.Key] = prop.Value;
                        }

                        foreach (var prop in valueProps)
                        {
                            if (ShouldSkipKey(prop.Key))
                                continue;

                            data[$"{prop.Key}"] = prop.Value;
                        }

                        items.Add(data);
                    }

                    var table = ToDataTable(items);
                    if (columns != null)
                    {
                        foreach (var column in columns)
                        {
                            if (!table.Columns.Contains(column.Trim()))
                                table.Columns.Add(column.Trim(), typeof(object));
                        }
                    }


                    table.TableName = field["testTitle"]?.ToString();

                    var existingTable = existingTables.FirstOrDefault(t => t.TableName == table.TableName);
                    if (existingTable != null)
                    {
                        foreach (DataRow row in table.Rows)
                        {
                            var newRow = existingTable.NewRow();
                            foreach (DataColumn col in existingTable.Columns)
                            {
                                newRow[col.ColumnName] = row[col.ColumnName];
                            }
                            existingTable.Rows.Add(newRow);
                        }

                        continue;
                    }

                    existingTables.Add(table);
                    yield return table;
                }
            }
        }
    }

    public class SheetProject : SheetProcessorBase, ISheetProcessor
    {

        public SheetProject(JsonNode? projectNode, JsonNode? boreholesNode) : base(projectNode, boreholesNode)
        {
        }

        public string Name => "Project";

        IEnumerable<DataTable> ISheetProcessor.Process()
        {
            if (_projectNode == null)
                yield break;

            var result = new List<Dictionary<string, object>>();
            var item = new Dictionary<string, object>();
            foreach (var node in _projectNode.AsObject())
            {
                if (node.Key == "extraTags")
                {
                    HandleExtraTags(item, node);
                    continue;
                }

                if (ShouldSkipKey(node.Key))
                {
                    continue;
                }


                item.Add(node.Key, node.Value);
            }

            result.Add(item);

            yield return ToDataTable(result);
        }

        private void HandleExtraTags(Dictionary<string, object> item, KeyValuePair<string, JsonNode?> node)
        {
            var tags = node.Value?.AsArray();
            if (tags == null)
                return;

            foreach (var tag in tags)
            {
                var name = tag?["name"]?.ToString();
                if (name == null || string.IsNullOrEmpty(name))
                    continue;

                var value = tag?["value"];

                item.Add($"{name}", value);
            }
        }
    }
}
