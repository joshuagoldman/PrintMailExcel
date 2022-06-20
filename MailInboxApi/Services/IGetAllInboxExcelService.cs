

using MailInboxApi.DataAccess;

namespace MailInboxApi.Services
{
    public interface IGetAllInboxExcelService
    {
        List<ExcelPrintInfo> GetAllExcelInfos(MailClient client);
    }
}