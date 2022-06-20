using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailInboxApi.DataAccess
{
    public readonly record struct ExcelPrintInfo(string Filename, string SheetName, string Info);
}
