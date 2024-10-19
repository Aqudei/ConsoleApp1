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
        {
        }

        public string Name => "Piezometers";

        public DataTable? Process() => ToDataTable(ProcessBoreHolesItems("piezometerData", "piezometer"));

    }

    public class SheetBackFill : SheetProcessorBase, ISheetProcessor
    {
        public SheetBackFill(JsonNode? projectNode, JsonNode? boreholesNode) : base(projectNode, boreholesNode)
        {
        }

        public string Name => "BackFill";

        public DataTable? Process() => ToDataTable(ProcessBoreHolesItems("piezometerData", "backfill"));
    }
    public class SheetDrillingDetails : SheetProcessorBase, ISheetProcessor
    {
        public SheetDrillingDetails(JsonNode? projectNode, JsonNode? boreholesNode) : base(projectNode, boreholesNode)
        {
        }

        public string Name => "Drilling Details";

        public DataTable? Process() => ToDataTable(ProcessBoreHolesItems("boringMethods"));
    }

    public class SheetStratigraphy : SheetProcessorBase, ISheetProcessor
    {
        public SheetStratigraphy(JsonNode? projectNode, JsonNode? boreholesNode) : base(projectNode, boreholesNode)
        {
        }

        public string Name => "Stratigraphy";

        public DataTable? Process() => ToDataTable(ProcessBoreHolesItems("stratigraphy"));
    }
    public class SheetDiscontinuities : SheetProcessorBase, ISheetProcessor
    {
        public SheetDiscontinuities(JsonNode? projectNode, JsonNode? boreholesNode) : base(projectNode, boreholesNode)
        {
        }

        public string Name => "Discontinuities";

        public DataTable? Process() => ToDataTable(ProcessBoreHolesItems("discontinuities"));
    }

    public class SheetDrillRuns : SheetProcessorBase, ISheetProcessor
    {
        public SheetDrillRuns(JsonNode? projectNode, JsonNode? boreholesNode) : base(projectNode, boreholesNode)
        {
        }

        public string Name => "Drill Runs";

        public DataTable? Process() => ToDataTable(ProcessBoreHolesItems("drillRuns"));
    }

    public class SheetSamples : SheetProcessorBase, ISheetProcessor
    {
        public SheetSamples(JsonNode? projectNode, JsonNode? boreholesNode) : base(projectNode, boreholesNode)
        {
        }

        public string Name => "Samples";

        public DataTable? Process()
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

                    data.Remove("typeIconUrl", out var _);

                    data["testHole"] = name;

                    var labTests = sample?["labTests"];

                    foreach (var property in labTests.AsObject())
                    {
                        foreach (var item in property.Value.AsArray())
                        {
                            var props = ExtractProperties(item);
                            foreach (var prop in props)
                            {
                                if (prop.Value is JsonValue)
                                    data[prop.Key] = prop.Value;
                            }
                        }
                    }

                    items.Add(data);
                }
            }

            return ToDataTable(items);
        }
    }
    public class SheetTestHole : SheetProcessorBase, ISheetProcessor
    {
        public string Name => "Test Holes";

        public SheetTestHole(JsonNode? projectNode, JsonNode? boreholesNode) : base(projectNode, boreholesNode)
        {
        }

        public DataTable? Process()
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
                                data[property.Key] = property.Value.GetValue<int>().ToString("N0");
                                continue;
                            }
                            data[property.Key] = property.Value;
                        }
                    }
                }

                items.Add(data);
            }

            return ToDataTable(items);
        }


    }



    public class SheetComments : SheetProcessorBase, ISheetProcessor
    {
        public SheetComments(JsonNode? projectNode, JsonNode? boreholesNode) : base(projectNode, boreholesNode)
        {
        }

        public string Name => "Comments";

        public DataTable? Process() => ToDataTable(ProcessBoreHolesItems("comments"));
    }

    public class SheetFieldTests : SheetProcessorBase
    {
        public SheetFieldTests(JsonNode? projectNode, JsonNode? boreholesNode) : base(projectNode, boreholesNode)
        {
        }

        public string Name => "Field Tests";

        public IEnumerable<DataTable> Process()
        {

            foreach (var borehole in _boreholesNode.AsArray())
            {
                var testHole = borehole["name"];

                var fieldTests = borehole["fieldTests"];
                foreach (var field in fieldTests.AsArray())
                {
                    var items = new List<Dictionary<string,object>>();

                    var fieldProps = ExtractProperties(field);

                    foreach (var val in field["values"].AsArray())
                    {
                        var valueProps = ExtractProperties(val);
                        var data = new Dictionary<string, object>();
                        data["testHole"] = testHole;

                        foreach (var prop in fieldProps)
                        {
                            if (ShouldSkipKey(prop.Key))
                                continue;
                            
                            //if (prop.Key == "testTitle")
                            //    continue;

                            data[prop.Key] = prop.Value;
                        }

                        foreach (var prop in valueProps)
                        {
                            if (ShouldSkipKey(prop.Key))
                                continue;

                            data[$"Test {prop.Key}"] = prop.Value;
                        }

                        items.Add(data);
                    }

                    var table = ToDataTable(items);
                    table.TableName = field["testTitle"].ToString();
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

        DataTable? ISheetProcessor.Process()
        {
            if (_projectNode == null)
                return null;

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

            return ToDataTable(result);
        }

        private void HandleExtraTags(Dictionary<string, object> item, KeyValuePair<string, JsonNode?> node)
        {
            var tags = node.Value.AsArray();
            foreach (var tag in tags)
            {
                var name = tag?["name"]?.ToString();
                if (name == null || string.IsNullOrEmpty(name))
                    continue;

                foreach (var n in tag.AsObject())
                {
                    if (n.Key == "value")
                    {
                        item.Add(name, n.Value);
                    }
                }
            }
        }
    }
}
