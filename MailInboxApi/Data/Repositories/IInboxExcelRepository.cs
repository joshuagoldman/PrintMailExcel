using MailInboxApi.DataAccess;

namespace MailInboxApi.Data.Repositories
{
    public interface IInboxExcelRepository
    {
        void SaveExcelRows(List<ExcelPrintInfo> client);
    }
}