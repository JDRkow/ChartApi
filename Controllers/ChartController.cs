using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ChartApi.Dto;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Caching.Memory;
using OfficeOpenXml;

namespace ChartApi.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class ChartController : ControllerBase
    {
        private IMemoryCache _cache;

        public ChartController(IMemoryCache cache)
        {
            _cache = cache;
        }

        [HttpGet]
        public List<TableDto> GetChartJson()
        {
            try
            {
                return _cache.Get<List<TableDto>>("excell-data");
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        [HttpPost]
        public IActionResult UploadTable([FromForm] IFormFile uploadTable)
        {
            using var uploadFileStream = uploadTable.OpenReadStream();
            var byteArrayFile = new byte[uploadTable.Length];
            uploadFileStream.Read(byteArrayFile, 0, (int)uploadTable.Length);

            var parsedTable = new List<TableDto>();

            try
            {
                using (MemoryStream stream = new MemoryStream(byteArrayFile))
                using (ExcelPackage excelPackage = new ExcelPackage(stream))
                {
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    var worksheet = excelPackage.Workbook.Worksheets.First();
                    for (var i = worksheet.Dimension.Start.Row + 1; i <= worksheet.Dimension.End.Row; i++)
                    {
                        parsedTable.Add(new TableDto
                        {
                            NameArea = worksheet.Cells[i, 1].Value.ToString(),
                            AreaParameter = int.Parse(worksheet.Cells[i, 2].Value.ToString())
                        });
                    }

                    _cache.Set("excell-data", parsedTable);
                }

                return Ok();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
    }
}
