using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PrintMailExcel
{
    public class ExcelPrint
    {
        public ExcelPrint(string filename, string sheetName, string content)
        {
            Filename = filename;
            SheetName = sheetName;
            Content = content;
        }

        public string Filename { get; }
        public string SheetName { get; }
        public string Content { get; }
    }
}
