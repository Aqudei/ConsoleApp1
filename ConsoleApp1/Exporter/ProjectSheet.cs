using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Text.Json.Nodes;
using System.Threading.Tasks;

namespace ConsoleApp1.Exporter
{
    internal class ProjectSheet : ExporterBase, ISheetProcessor
    {
        private JsonNode? _projectNode;
        private JsonNode? _boreholesNode;

        public ProjectSheet(JsonNode? projectNode, JsonNode? boreholesNode)
        {
            _projectNode = projectNode;
            _boreholesNode = boreholesNode;
        }

        public string Name => "Project";

        DataTable ISheetProcessor.Process()
        {
            var result = new List<Dictionary<string, object>>();
            var item = new Dictionary<string, object>();
            foreach (var node in _projectNode.AsObject())
            {
                if (node.Key == "extraTags")
                {
                    continue;
                }

                if (ShouldSkipNode(node))
                {
                    continue;
                }

                item.Add(node.Key, node.Value);
            }

            result.Add(item);

            return ToDataTable(result);
        }

        private bool ShouldSkipNode(KeyValuePair<string, JsonNode?> node)
        {
            return node.Key.EndsWith("Id") || node.Key == "id";
        }
    }
}
