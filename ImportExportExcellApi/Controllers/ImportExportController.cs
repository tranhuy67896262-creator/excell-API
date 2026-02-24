using ImportExportExcellApi.Data;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Http.HttpResults;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.Style;
using System.Drawing;

namespace ImportExportExcellApi.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ImportExportController : ControllerBase
    {
        private readonly IWebHostEnvironment _environment;

        // Cấu hình đường dẫn
        // Giả sử folder Templates nằm cùng cấp với file .csproj hoặc trong root dự án
        private readonly string _templateFolderName = "Templates";
        private readonly string _exportFolderName = "Exports";
        private readonly string _baseTemplateName = "Employee_Template.xlsx"; // Tên file base bạn đã format sẵn

        public ImportExportController(IWebHostEnvironment environment)
        {
            _environment = environment;
            // Khởi tạo data mẫu nếu chưa có
            AppDataContext.Initialize();
        }
        //viết 1 func dẽ đọc file excell base và sẽ thực hiện add công thức và gen ra file tên là tên file + realase in ptoduct và đọc file đó fill data
        // sau đó mới export ra file chứa data mẫu chứa các dropdow , các dropdown khi import sẽ lấy được là ID , add công thức các thứ các cột , có cột hơi đặc biệt với tôi là file

        [HttpGet("export-to-product")]
        public IActionResult ExportToProduct()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // 1. Xác định đường dẫn file Template
            // Cách 1: Nếu folder Templates nằm trong thư mục gốc của dự án (khi run sẽ copy sang bin)
            string templatePath = Path.Combine(_environment.ContentRootPath, _templateFolderName, _baseTemplateName);

            // Cách 2: Nếu bạn để folder Templates ngang hàng với folder bin/obj (bên ngoài ContentRoot)
            // string templatePath = Path.Combine(Directory.GetParent(_environment.ContentRootPath)?.FullName ?? "", _templateFolderName, _baseTemplateName);

            if (!System.IO.File.Exists(templatePath))
            {
                return NotFound(new { message = $"Không tìm thấy file template tại: {templatePath}. Vui lòng đảm bảo file '{_baseTemplateName}' tồn tại trong folder '{_templateFolderName}'." });
            }

            return Ok(templatePath);
        }
    }
}
