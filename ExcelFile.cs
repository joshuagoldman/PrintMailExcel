using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PrintMailExcel
{
    public class ExcelFile
    {
        public ExcelFile(List<ExcelSheet> sheets, string excelFileName, Guid fileGuid)
        {
            Sheets = sheets;
            ExcelFileName = excelFileName;
            FileGuid = fileGuid;
        }

        public List<ExcelSheet> Sheets { get; }
        public string ExcelFileName { get; }
        public Guid FileGuid { get; }
    }
}
