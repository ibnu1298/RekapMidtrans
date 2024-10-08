using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using RekapMidtrans.DTO;
using RekapMidtrans.Service;

namespace RekapMidtrans.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class RekapController : ControllerBase
    {
        private readonly IExcel _excel;
        public RekapController(IExcel excel)
        {
            _excel = excel;
        }
        [HttpPost]
        public async Task<ActionResult> RekapMidtrans([FromForm] UploadExcelRequest request)
        {
            try
            {
                DownloadExcelResponse? response = null;
                if (Request.HasFormContentType)
                {
                    var requestProperties = _excel.ExtractUploadRequest(Request.Form.Files);
                    requestProperties.token = request.token;
                    requestProperties.URLgetIdOrder = request.URLgetIdOrder;
                    requestProperties.URLgetOrderDetail = request.URLgetOrderDetail;

                    response = await _excel.DownloadRekapMidtrans(requestProperties);
                }
                if (response == null)
                {
                    return BadRequest("Unknown type");
                }
                else
                {
                    return File(response.FileContents, response.ContentType, response.FileDownloadName);
                }
            }
            catch (Exception ex)
            {
                return Problem(ex.Message);
            }
        }
    }
}
