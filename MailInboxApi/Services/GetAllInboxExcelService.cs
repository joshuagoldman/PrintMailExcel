using MailInboxApi.Data.Repositories;
using MailInboxApi.DataAccess;
using PrintMailExcel;
using MailClient = MailInboxApi.DataAccess.MailClient;

namespace MailInboxApi.Services
{
    public class GetAllInboxExcelService : IGetAllInboxExcelService
    {
        private readonly IInboxExcelRepository _inboxExcelRepository;
        private readonly IConfiguration _configuration;
        public GetAllInboxExcelService(IInboxExcelRepository inboxExcelRepository,
                                       IConfiguration configuration)
        {
            _inboxExcelRepository = inboxExcelRepository;
            _configuration = configuration;
        }

        public List<ExcelPrintInfo> GetAllExcelInfos(MailClient client)
        {
            PrintMailExcel.MailClient mailHandling = new PrintMailExcel.MailClient(client.EmailAddress,
                                                                                   client.Password);

            List<ExcelFile> excelInfos = new List<ExcelFile>();
            if (mailHandling.TryGetExcelFilesInfo(_configuration["Storage:ExcelFiles"], out excelInfos))
            {
                List<ExcelPrint> excelPrintInfos = mailHandling.GetRows(excelInfos);
                IEnumerable<ExcelPrintInfo> excelInfosDb = excelPrintInfos.Select(info => new ExcelPrintInfo
                {
                    SheetName = info.SheetName,
                    Filename = info.Filename,
                    Info = info.Content
                });

                _inboxExcelRepository.SaveExcelRows(excelInfosDb.ToList());
                PrintMailExcel.MailClient.PrintExcelFilesinfoRows(excelInfos);

                return excelInfosDb.ToList();
            }
            else
                throw new Exception("Could not fetch any excel file info from inbox!");
        }
    }
}
