using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp1.Exporter
{
    public abstract class ExporterBase
    {
        public DataTable ToDataTable(List<Dictionary<string, object>> data)
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
