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

            foreach (var boreHole in _boreholesNode?.AsArray() ?? new JsonArray())
            {
                var targetData = string.IsNullOrEmpty(nestedKey) ? boreHole?[arrayKey] : boreHole?[arrayKey]?[nestedKey];
                if (targetData == null) continue;

                foreach (var dataNode in targetData.AsArray())
                {
                    var data = ExtractProperties(dataNode);
                    data["testHole"] = boreHole["name"].AsValue();
                    items.Add(data);
                }
            }

            return items;
        }

        protected Dictionary<string, object> ExtractProperties(JsonNode? dataNode)
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
            if (data == null || data.Count == 0)
                return new DataTable();

            var dataTable = PrepareDataTable(data);

            
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

        private DataTable PrepareDataTable(List<Dictionary<string, object>> data)
        {
            var dataTable = new DataTable();

            // Get all unique keys (columns) from the dictionaries
            HashSet<string> columns = new HashSet<string>();
            foreach (var dict in data)
            {
                foreach (var key in dict.Keys)
                {
                    columns.Add(key);
                }
            }

            // Create columns for the DataTable
            foreach (var column in columns)
            {
                dataTable.Columns.Add(column, typeof(object)); // Use object type for flexibility
            }


            // Column name to move to first position
            string columnNameToMove = "testHole";

            // Check if the column exists and move it to the first position
            if (dataTable.Columns.Contains(columnNameToMove))
            {
                dataTable = MoveColumnToFirst(dataTable, columnNameToMove);
            }

            return dataTable;
        }

        private DataTable MoveColumnToFirst(DataTable table, string columnName)
        {
            // Create a new DataTable with the desired order
            DataTable newTable = new DataTable();

            // Add the specified column to the new table as the first column
            newTable.Columns.Add(columnName, table.Columns[columnName].DataType);

            // Add the rest of the columns (excluding the one being moved)
            foreach (DataColumn column in table.Columns)
            {
                if (column.ColumnName != columnName)
                {
                    newTable.Columns.Add(column.ColumnName, column.DataType);
                }
            }

            // Copy the data row by row
            foreach (DataRow row in table.Rows)
            {
                DataRow newRow = newTable.NewRow();
                foreach (DataColumn column in table.Columns)
                {
                    newRow[column.ColumnName] = row[column.ColumnName];
                }
                newTable.Rows.Add(newRow);
            }

            return newTable;
        }
    }
}
