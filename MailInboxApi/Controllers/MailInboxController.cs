using MailInboxApi.DataAccess;
using MailInboxApi.Services;
using Microsoft.AspNetCore.Mvc;

namespace MailInboxApi.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class MailInboxController : Controller
    {
        IGetAllInboxExcelService _getAllInboxExcelService;
        public MailInboxController(IGetAllInboxExcelService getAllInboxExcelService)
        {
            _getAllInboxExcelService = getAllInboxExcelService;
        }

        [HttpPost("/ExcelFiles")]
        public IActionResult GetAllInboxExcelFilesInfo(MailClient client)
        {
            try
            {
                List<ExcelPrintInfo> excelPrintInfo = _getAllInboxExcelService.GetAllExcelInfos(client);

                return Json(excelPrintInfo);
            }
            catch (Exception e)
            {

                return BadRequest(e.ToString());
            }
        }
    }
}
