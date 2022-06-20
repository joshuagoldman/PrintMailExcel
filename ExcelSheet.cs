using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PrintMailExcel
{
    public class ExcelSheet
    {
        public ExcelSheet(List<List<string>> table, string sheetName)
        {
            Table = table;
            SheetName = sheetName;
        }
        public List<List<string>> Table { get; }
        public string SheetName { get; }
    }
}
