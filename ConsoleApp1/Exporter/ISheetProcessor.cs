using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Text.Json.Nodes;
using System.Threading.Tasks;

namespace ConsoleApp1.Exporter
{
    internal interface ISheetProcessor
    {
        string Name { get;}

        void FormatSheet(IXLWorksheet worksheet);
        IEnumerable<DataTable> Process();
    }
}
