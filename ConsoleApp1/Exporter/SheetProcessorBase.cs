using ClosedXML.Excel;
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

        protected virtual Dictionary<string, string>? GetColumnsMapping() => null;

        protected SheetProcessorBase(JsonNode? projectNode, JsonNode? boreholesNode)
        {
            _projectNode = projectNode;
            _boreholesNode = boreholesNode;


            _unitSystem = projectNode?["unitSystem"]?.ToString();
            _crsMeasurementUnit = projectNode?["crsMeasurementUnit"]?.ToString();
            _unit = _unitSystem != null && _unitSystem == "Imperial" ? "ft" : "m";
        }
        protected void ResizeAndWrapColumn(IXLWorksheet worksheet, string columnName, double width)
        {
            // Find the column number based on the header name (assumes headers are in the first row)
            var headerRow = worksheet.Row(1);
            var column = headerRow.Cells().FirstOrDefault(c => c.Value.ToString() == columnName);

            if (column != null)
            {
                int columnNumber = column.Address.ColumnNumber;

                // Set the width of the column
                worksheet.Column(columnNumber).Width = width;

                // Enable text wrapping for the entire column
                worksheet.Column(columnNumber).Style.Alignment.WrapText = true;
                worksheet.Row(column.Address.RowNumber).AdjustToContents();
            }
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
                    if (GetIgnoredProperties().Select(p=>p.ToLower()).Contains(property.Key.ToLower())) 
                        continue;

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
                            if (GetIgnoredProperties().Select(p => p.ToLower()).Contains(property.Key.ToLower()))
                                continue;
                            data[prop.Key] = prop.Value;
                        }
                    }
                }
            }
            return data;
        }

        protected DataTable ToDataTable(IEnumerable<Dictionary<string, object>> data)
        {
            if (data == null || data.Count() == 0)
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

            return ReOrderColumns(PostProcessDataTable(dataTable));
        }
        protected virtual DataTable PostProcessDataTable(DataTable dataTable)
        {
            var dt = SetFirstColumn(dataTable, "testHole");

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

                var columnMappings = GetColumnsMapping();

                if (columnMappings != null && columnMappings.TryGetValue(newColumnName, out var desiredName))
                {
                    newColumnName = desiredName;
                }

                column.ColumnName = newColumnName;
            }
            return dt;
        }

        protected virtual Dictionary<string, int> GetColumnsOrdering()
        {
            return new Dictionary<string, int>();

        }
        protected DataTable ReOrderColumns(DataTable table)
        {
            var newTable = new DataTable();
            // var columns = new Dictionary<string, int> { { "Test Hole", 0 }, { "Test Title", 1 }, { "Depth (ft)", 2 }, { "Value", 3 } };

            // Step 1: Add columns from the dictionary to the new table in the specified order
            foreach (var column in GetColumnsOrdering().OrderBy(c => c.Value))
            {
                if (table.Columns.Contains(column.Key))
                {
                    newTable.Columns.Add(column.Key, table.Columns[column.Key].DataType);
                }
            }

            // Step 2: Add any remaining columns that are not in the dictionary (in their original order)
            foreach (DataColumn col in table.Columns)
            {
                if (!newTable.Columns.Contains(col.ColumnName))
                {
                    newTable.Columns.Add(col.ColumnName, col.DataType);
                }
            }

            // Step 3: Copy data from the original table to the new table
            foreach (DataRow row in table.Rows)
            {
                var newRow = newTable.NewRow();

                // Copy values based on matching column names
                foreach (DataColumn col in newTable.Columns)
                {
                    newRow[col.ColumnName] = row[col.ColumnName];
                }

                newTable.Rows.Add(newRow);
            }

            return newTable;
        }

        private string AddSpacesToSentence(string input)
        {
            // Use regular expression to insert spaces before each capital letter
            string spacedString = Regex.Replace(input, "(\\B[A-Z])", " $1");

            // Capitalize the first letter of the sentence
            return char.ToUpper(spacedString[0]) + spacedString.Substring(1);
        }

        private DataTable PrepareDataTable(IEnumerable<Dictionary<string, object>> data)
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

        private DataTable SetFirstColumn(DataTable table, string columnName)
        {
            if (!table.Columns.Contains(columnName))
            {
                return table;
            }

            // Create a new DataTable with the desired order
            var newTable = new DataTable();

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

        public virtual void FormatSheet(IXLWorksheet worksheet)
        {

        }
    }
}
