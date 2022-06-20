using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PrintMailExcel
{
    internal class ExcelHandling
    {
        internal ExcelHandling()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        internal List<ExcelSheet> GetExcelTables(string filePath)
        {
            var workSheetsTable = new List<ExcelSheet>();
            using (var excelPckg = new ExcelPackage(filePath))
            {
                var workSheets = excelPckg.Workbook.Worksheets;
                for (int i = 0; i < workSheets.Count; i++)
                {
                    var rowCount = workSheets[i].Dimension.End.Row;
                    var colCount = workSheets[i].Dimension.End.Column;
                    var workSheetTable = new List<List<string>>();

                    for (int j = 1; j <= rowCount; j++)
                    {
                        var tempRow = new List<string>();
                        for (int k = 1; k <= colCount; k++)
                        {
                            var tempVal = workSheets[i].Cells[j, k].Value;
                            tempRow.Add(tempVal == null ? "" : tempVal.ToString());
                        }
                        workSheetTable.Add(tempRow);
                    }
                    var tempWorkSheet = new ExcelSheet(workSheetTable, workSheets[i].Name);
                    workSheetsTable.Add(tempWorkSheet);
                }
            }

            return workSheetsTable;
        }
    }
}
