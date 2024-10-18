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
                    HandleExtraTags(item, node);
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

        private bool ShouldSkipNode(KeyValuePair<string, JsonNode?> node)
        {
            return node.Key.EndsWith("Id") || node.Key == "id";
        }
    }
}
