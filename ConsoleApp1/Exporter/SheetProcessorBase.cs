using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ConsoleApp1.Exporter
{
    public abstract class SheetProcessorBase
    {
        protected JsonNode? _projectNode;
        protected JsonNode? _boreholesNode;
        protected string? _unitSystem;
        protected string? _crsMeasurementUnit;
        protected string _unit;

        protected SheetProcessorBase(JsonNode? projectNode, JsonNode? boreholesNode)
        {
            _projectNode = projectNode;
            _boreholesNode = boreholesNode;


            _unitSystem = projectNode?["unitSystem"]?.ToString();
            _crsMeasurementUnit = projectNode?["crsMeasurementUnit"]?.ToString();
            _unit = _unitSystem != null && _unitSystem == "Imperial" ? "ft" : "m";
        }

        protected bool ShouldSkipKey(string key)
        {
            return key.EndsWith("Id") || key == "id" || string.IsNullOrWhiteSpace(key);
        }

        protected List<Dictionary<string, object>> ProcessBoreHolesItems(string arrayKey, string nestedKey = null, IEnumerable<string> include = null)
        {
            var items = new List<Dictionary<string, object>>();

            foreach (var boreHole in _boreholesNode?.AsArray() ?? new JsonArray())
            {
                var targetData = string.IsNullOrEmpty(nestedKey) ? boreHole?[arrayKey] : boreHole?[arrayKey]?[nestedKey];
                if (targetData == null) continue;

                foreach (var dataNode in targetData.AsArray())
                {
                    var data = ExtractProperties(dataNode, include);
                    data["testHole"] = boreHole["name"].AsValue();

                    foreach (var item in GetIgnoredProperties())
                    {
                        data.Remove(item, out var _);
                    }

                    items.Add(data);
                }
            }

            return items;
        }
        protected virtual IEnumerable<string> GetIgnoredProperties()
        {
            return new List<string>();
        }

        protected Dictionary<string, object> ExtractProperties(JsonNode? dataNode, IEnumerable<string> include = null)
        {
            var data = new Dictionary<string, object>();
            foreach (var property in dataNode.AsObject())
            {
                if (ShouldSkipKey(property.Key)) continue;
                if (property.Value is JsonValue || property.Value is null)
                {
                    data[property.Key] = property.Value;
                }
                else if (property.Value is JsonObject)
                {
                    if (include != null && !include.Contains(property.Key))
                        continue;

                    foreach (var prop in property.Value.AsObject())
                    {
                        if (ShouldSkipKey(prop.Key)) continue;

                        if (prop.Value is JsonValue || prop.Value is null)
                        {
                            data[prop.Key] = prop.Value;
                        }
                    }
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


            return PostProcessDataTable(dataTable);
        }

        private DataTable PostProcessDataTable(DataTable dataTable)
        {
            var dt = MoveColumnToFirst(dataTable, "testHole");

            foreach (DataColumn column in dt.Columns)
            {
                var newColumnName = AddSpacesToSentence(column.ColumnName);
                if (newColumnName.EndsWith("Depth"))
                {
                    newColumnName = $"{newColumnName} ({_unit})";
                }

                if (newColumnName.EndsWith("Elev") || newColumnName.EndsWith("Elevation"))
                {
                    newColumnName = $"{newColumnName} ({_crsMeasurementUnit})";
                }

                column.ColumnName = newColumnName;
            }
            return dt;
        }

        private string AddSpacesToSentence(string input)
        {
            // Use regular expression to insert spaces before each capital letter
            string spacedString = Regex.Replace(input, "(\\B[A-Z])", " $1");

            // Capitalize the first letter of the sentence
            return char.ToUpper(spacedString[0]) + spacedString.Substring(1);
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

            return dataTable;
        }

        private DataTable MoveColumnToFirst(DataTable table, string columnName)
        {
            if (!table.Columns.Contains(columnName))
            {
                return table;
            }

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
