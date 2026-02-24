using ImportExportExcellApi.Data;
using ImportExportExcellApi.Entities;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.Style;
using System.Drawing;
using System.IO;
using System.Linq;

namespace ImportExportExcellApi.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ImportExportController : ControllerBase
    {
        private readonly IWebHostEnvironment _environment;
        private readonly string _templateFolderName = "Templates";
        private readonly string _baseTemplateName = "Employee_Template_Base.xlsx";
        private readonly string _releaseTemplateName = "Employee_Template_Release.xlsx";

        public ImportExportController(IWebHostEnvironment environment)
        {
            _environment = environment;
            AppDataContext.Initialize();

            string templatePath = Path.Combine(_environment.ContentRootPath, _templateFolderName);
            if (!Directory.Exists(templatePath))
            {
                Directory.CreateDirectory(templatePath);
            }
        }

        [HttpGet("gen-excel-template")]
        public IActionResult GenExcelTemplate()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            string templateFolder = Path.Combine(_environment.ContentRootPath, _templateFolderName);
            string baseTemplatePath = Path.Combine(templateFolder, _baseTemplateName);

            if (!System.IO.File.Exists(baseTemplatePath))
            {
                return NotFound(new { message = $"Không tìm thấy file Base Template tại: {baseTemplatePath}" });
            }

            byte[] fileBytes;
            string fileName = $"Mau_Nhap_Lieu_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";

            using (var package = new ExcelPackage(new FileInfo(baseTemplatePath)))
            {
                // ========================================
                // 1. TẠO SHEET ẨN CHỨA DỮ LIỆU TRA CỨU
                // ========================================
                string lookupSheetName = "Data_Lookup";

                // Xóa sheet cũ nếu tồn tại
                if (package.Workbook.Worksheets.Any(w => w.Name == lookupSheetName))
                {
                    package.Workbook.Worksheets.Delete(lookupSheetName);
                }

                var lookupWs = package.Workbook.Worksheets.Add(lookupSheetName);
                lookupWs.Hidden = eWorkSheetHidden.Hidden; // Ẩn sheet

                // Header cho sheet lookup
                lookupWs.Cells[1, 1].Value = "CODE";
                lookupWs.Cells[1, 2].Value = "FULL_NAME";
                lookupWs.Cells[1, 1, 1, 2].Style.Font.Bold = true;

                // Đổ dữ liệu từ DB vào sheet lookup
                var lookupData = AppDataContext.Employees
                    .Join(AppDataContext.EmployeeCvs,
                        emp => emp.EmployeeId,
                        cv => cv.Id,
                        (emp, cv) => new { emp.Code, cv.FullName })
                    .OrderBy(x => x.Code)
                    .ToList();

                for (int i = 0; i < lookupData.Count; i++)
                {
                    lookupWs.Cells[i + 2, 1].Value = lookupData[i].Code;
                    lookupWs.Cells[i + 2, 2].Value = lookupData[i].FullName;
                }

                // Tạo Named Range cho vùng dữ liệu lookup (Code + FullName)
                string lookupRangeName = "DanhSachNhanVien";
                package.Workbook.Names.Add(lookupRangeName,
                    lookupWs.Cells[$"A2:B{lookupData.Count + 1}"]);

                // Tạo Named Range chỉ chứa Code cho dropdown
                string codeListRangeName = "DanhSachCode";
                package.Workbook.Names.Add(codeListRangeName,
                    lookupWs.Cells[$"A2:A{lookupData.Count + 1}"]);

                // ========================================
                // 2. XỬ LÝ SHEET CHÍNH - GIỮ NGUYÊN CẤU TRÚC
                // ========================================
                var ws = package.Workbook.Worksheets[0]; // Sheet đầu tiên

                // ✅ Không ghi đè header (row 4) và hướng dẫn (row 5)
                // ✅ Chỉ áp dụng dropdown và formula từ row 6 trở đi
                int startRow = 6;
                int endRow = 100; // Số dòng tối đa cho phép nhập (có thể điều chỉnh)

                // ========================================
                // 3. ÁP DỤNG DROPDOWN CHO CỘT A (MÃ NHÂN VIÊN)
                // ========================================
                // Cột A = column 1 (Mã nhân viên)
                ApplyDropdown(ws, startRow, endRow, 1, codeListRangeName);

                // ========================================
                // 4. ÁP DỤNG VLOOKUP CHO CỘT B (HỌ VÀ TÊN)
                // ========================================
                // Cột B = column 2 (Họ và tên) - tự động hiện khi nhập mã ở cột A
                for (int row = startRow; row <= endRow; row++)
                {
                    // Công thức: Nếu cột A rỗng thì để trống, ngược lại tra cứu tên
                    string vlookupFormula = $"=IF(A{row}=\"\", \"\", VLOOKUP(A{row}, DanhSachNhanVien, 2, FALSE))";
                    ws.Cells[row, 2].Formula = vlookupFormula;

                    // Format cột Họ tên: màu xám nhạt để biết là tự động
                    ws.Cells[row, 2].Style.Font.Color.SetColor(Color.DarkGray);
                    ws.Cells[row, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells[row, 2].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(245, 245, 245));
                }

                // ========================================
                // 5. ÁP DỤNG DROPDOWN CHO CÁC CỘT KHÁC (nếu có)
                // ========================================
                // Cột F = column 6 (Đơn vị đào tạo) - nếu có Named Range
                if (ws.Workbook.Names.Any(n => n.Name.Equals("DanhSachBoPhan", StringComparison.OrdinalIgnoreCase)))
                {
                    ApplyDropdown(ws, startRow, endRow, 6, "DanhSachBoPhan");
                }

                // ========================================
                // 6. FORMAT & HOÀN TẤT
                // ========================================

                // Border cho vùng dữ liệu (từ row 6 đến row 100)
                // Giả sử template có 7 cột (A đến G)
                var dataRange = ws.Cells[$"A{startRow}:G{endRow}"];
                dataRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                dataRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                dataRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                dataRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;

                // Căn giữa cho cột STT (nếu có)
                ws.Column(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                fileBytes = package.GetAsByteArray();
            }

            return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        }

        /// <summary>
        /// Áp dụng Data Validation dạng Dropdown List cho một vùng ô
        /// </summary>
        private void ApplyDropdown(ExcelWorksheet ws, int startRow, int endRow, int column, string namedRange)
        {
            // Kiểm tra Named Range có tồn tại trong Workbook
            if (ws.Workbook.Names.Any(n => n.Name.Equals(namedRange, StringComparison.OrdinalIgnoreCase)))
            {
                var range = ws.Cells[startRow, column, endRow, column];
                var validation = range.DataValidation.AddListDataValidation();

                // Dùng INDIRECT để tham chiếu đến Named Range
                validation.Formula.ExcelFormula = $"INDIRECT(\"{namedRange}\")";

                validation.ShowErrorMessage = true;
                validation.ErrorStyle = ExcelDataValidationWarningStyle.warning;
                validation.ErrorTitle = "Giá trị không hợp lệ";
                validation.Error = "Vui lòng chọn mã có trong danh sách dropdown.";
                validation.ShowInputMessage = true;
                validation.PromptTitle = "💡 Hướng dẫn";
                validation.Prompt = "Chọn từ danh sách hoặc paste mã vào, tên sẽ tự động hiện ra.";
            }
        }
    }
}