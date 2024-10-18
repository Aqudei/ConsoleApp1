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

    public class SheetTestHole : SheetProcessorBase, ISheetProcessor
    {
        public string Name => "Test Hole";

        public SheetTestHole(JsonNode? projectNode, JsonNode? boreholesNode) : base(projectNode, boreholesNode)
        {
        }

        public DataTable? Process()
        {
            var items = new List<Dictionary<string, object>>();

            foreach (var item in _boreholesNode.AsArray())
            {
                var data = new Dictionary<string, object>();
                foreach (var property in item.AsObject())
                {
                    if (ShouldSkipKey(property.Key))
                        continue;

                    if (property.Key == "drillingGroundwaterLevels")
                    {
                        data["groundwaterDepth"] = property.Value["groundwaterDepth"];
                        data["groundwaterElev"] = property.Value["groundwaterElev"];
                    }

                    if (property.Key == "progressStatus")
                    {
                        data["progressStatus"] = property.Value["name"];
                    }

                    if (property.Key == "completionNotes")
                    {
                        foreach (var completionNotesNode in property.Value.AsObject())
                        {
                            data[completionNotesNode.Key] = completionNotesNode.Value;
                            data[completionNotesNode.Key] = completionNotesNode.Value;
                        }
                    }

                    if (property.Key == "sptHammer")
                    {
                        foreach (var sptHammerNode in property.Value.AsObject())
                        {
                            data[$"SPT{sptHammerNode.Key}"] = sptHammerNode.Value;
                            data[$"SPT{sptHammerNode.Key}"] = sptHammerNode.Value;
                        }
                    }

                    if (property.Value is JsonValue jsonValue)
                    {
                        data[property.Key] = jsonValue;
                    }


                    if (property.Value is null)
                    {
                        data[property.Key] = property.Value;
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

    public class SheetFieldTests : SheetProcessorBase, ISheetProcessor
    {
        public SheetFieldTests(JsonNode? projectNode, JsonNode? boreholesNode) : base(projectNode, boreholesNode)
        {
        }

        public string Name => "Field Tests";

        public DataTable? Process() => ToDataTable(ProcessBoreHolesItems("fieldTests"));

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
