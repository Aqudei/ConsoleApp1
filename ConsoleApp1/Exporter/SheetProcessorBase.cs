using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Text.Json.Nodes;
using System.Threading.Tasks;

namespace ConsoleApp1.Exporter
{
    public abstract class SheetProcessorBase
    {
        protected JsonNode? _projectNode;
        protected JsonNode? _boreholesNode;

        protected SheetProcessorBase(JsonNode? projectNode, JsonNode? boreholesNode)
        {
            _projectNode = projectNode;
            _boreholesNode = boreholesNode;
        }

        protected bool ShouldSkipKey(string key)
        {
            return key.EndsWith("Id") || key == "id";
        }

        protected List<Dictionary<string, object>> ProcessBoreHolesItems(string arrayKey, string nestedKey = null)
        {
            var items = new List<Dictionary<string, object>>();

            foreach (var item in _boreholesNode?.AsArray() ?? new JsonArray())
            {
                var targetData = string.IsNullOrEmpty(nestedKey) ? item?[arrayKey] : item?[arrayKey]?[nestedKey];
                if (targetData == null) continue;

                foreach (var dataNode in targetData.AsArray())
                {
                    var data = ExtractProperties(dataNode);
                    items.Add(data);
                }
            }

            return items;
        }

        private Dictionary<string, object> ExtractProperties(JsonNode? dataNode)
        {
            var data = new Dictionary<string, object>();
            foreach (var property in dataNode.AsObject())
            {
                if (ShouldSkipKey(property.Key)) continue;
                if (property.Value is JsonValue || property.Value is null)
                {
                    data[property.Key] = property.Value;
                }
            }
            return data;
        }

        protected DataTable ToDataTable(List<Dictionary<string, object>> data)
        {
            var dataTable = new DataTable();

            if (data == null || data.Count == 0)
                return dataTable;

            // Add columns to the DataTable based on keys from the first dictionary
            foreach (var key in data[0].Keys)
            {
                dataTable.Columns.Add(key, typeof(object));
            }

            // Add rows to the DataTable
            foreach (var dict in data)
            {
                var row = dataTable.NewRow();
                foreach (var kvp in dict)
                {
                    row[kvp.Key] = kvp.Value ?? DBNull.Value;
                }
                dataTable.Rows.Add(row);
            }

            return dataTable;
        }

    }
}
