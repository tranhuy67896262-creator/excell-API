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
                if (package.Workbook.Worksheets.Any(w => w.Name == lookupSheetName))
                {
                    package.Workbook.Worksheets.Delete(lookupSheetName);
                }

                var lookupWs = package.Workbook.Worksheets.Add(lookupSheetName);
                lookupWs.Hidden = eWorkSheetHidden.Hidden;

                lookupWs.Cells[1, 1].Value = "CODE";
                lookupWs.Cells[1, 2].Value = "FULL_NAME";
                lookupWs.Cells[1, 3].Value = "EMPLOYEE_ID";
                lookupWs.Cells[1, 1, 1, 3].Style.Font.Bold = true;

                var lookupData = AppDataContext.Employees
                    .Join(AppDataContext.EmployeeCvs,
                        emp => emp.EmployeeId,
                        cv => cv.Id,
                        (emp, cv) => new { emp.Code, emp.Id, cv.FullName })
                    .OrderBy(x => x.Code)
                    .ToList();

                for (int i = 0; i < lookupData.Count; i++)
                {
                    lookupWs.Cells[i + 2, 1].Value = lookupData[i].Code;
                    lookupWs.Cells[i + 2, 2].Value = lookupData[i].FullName;
                    lookupWs.Cells[i + 2, 3].Value = lookupData[i].Id;
                }

                // Tạo Named Ranges
                package.Workbook.Names.Add("DanhSachNhanVien", lookupWs.Cells[$"A2:B{lookupData.Count + 1}"]);
                package.Workbook.Names.Add("DanhSachCodeId", lookupWs.Cells[$"A2:C{lookupData.Count + 1}"]);
                package.Workbook.Names.Add("DanhSachCode", lookupWs.Cells[$"A2:A{lookupData.Count + 1}"]);

                // ========================================
                // 2. XỬ LÝ SHEET CHÍNH - GIỮ NGUYÊN FORMAT
                // ========================================
                var ws = package.Workbook.Worksheets[0];
                int startRow = 6;
                int endRow = 100;

                // ✅ 1. Áp dụng Dropdown (Data Validation) - Không ảnh hưởng visual
                ApplyDropdown(ws, startRow, endRow, 1, "DanhSachCode");
                if (ws.Workbook.Names.Any(n => n.Name.Equals("DanhSachBoPhan", StringComparison.OrdinalIgnoreCase)))
                {
                    ApplyDropdown(ws, startRow, endRow, 6, "DanhSachBoPhan");
                }

                // ✅ 2. Đổ công thức và giá trị - KHÔNG GÁN STYLE LẠI
                for (int row = startRow; row <= endRow; row++)
                {
                    // Cột B: FullName (VLOOKUP)
                    string nameFormula = $"=IF(A{row}=\"\", \"\", VLOOKUP(A{row}, DanhSachNhanVien, 2, FALSE))";
                    var nameCell = ws.Cells[row, 2];
                    nameCell.Formula = nameFormula;

                    // Cột Z (26): Hidden ID
                    string idFormula = $"=IF(A{row}=\"\", \"\", VLOOKUP(A{row}, DanhSachCodeId, 3, FALSE))";
                    ws.Cells[row, 26].Formula = idFormula;
                }

                // 🔐 Ẩn cột ID (Cột Z)
                ws.Column(26).Hidden = true;
                ws.Column(26).Width = 0.1;

                
                var dataRange = ws.Cells[$"A{startRow}:Y{endRow}"];
                dataRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                dataRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                dataRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                dataRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;

                fileBytes = package.GetAsByteArray();
            }

            return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        }

        [HttpPost("import-employee")]
        public IActionResult ImportEmployee(IFormFile file)
        {
            if (file == null || file.Length == 0)
            {
                return BadRequest(new { message = "Vui lòng chọn file Excel để import." });
            }

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var employees = new List<EmployeeImportResult>();
            var errors = new List<string>();

            using (var stream = new MemoryStream())
            {
                file.CopyTo(stream);
                using (var package = new ExcelPackage(stream))
                {
                    var ws = package.Workbook.Worksheets[0];

                    int startRow = 6;
                    int row = startRow;
                    int codeColumn = 1;        // Cột A: Code
                    int hiddenIdColumn = 26;   // ✅ Cột Z: EmployeeId (ẩn)

                    while (ws.Cells[row, codeColumn].Value != null)
                    {
                        var code = ws.Cells[row, codeColumn].Value?.ToString()?.Trim();

                        if (!string.IsNullOrEmpty(code))
                        {
                            // ✅ Đọc EmployeeId từ cột ẩn (cột Z)
                            var idValue = ws.Cells[row, hiddenIdColumn].Value;

                            if (idValue != null && long.TryParse(idValue.ToString(), out long employeeId))
                            {
                                // ✅ Kiểm tra lại ID có tồn tại trong DB (bảo mật)
                                var employee = AppDataContext.Employees
                                    .FirstOrDefault(e => e.Id == employeeId && e.Code == code);

                                if (employee != null)
                                {
                                    employees.Add(new EmployeeImportResult
                                    {
                                        RowNumber = row,
                                        Code = code,
                                        EmployeeId = employeeId,              // ✅ Lấy trực tiếp từ file
                                        EmployeeCvId = employee.EmployeeId,
                                        FullName = AppDataContext.EmployeeCvs
                                            .FirstOrDefault(cv => cv.Id == employee.EmployeeId)?.FullName
                                    });
                                }
                                else
                                {
                                    errors.Add($"Dòng {row}: Dữ liệu không khớp (Code: {code}, ID: {employeeId})");
                                }
                            }
                            else
                            {
                                // Fallback: Nếu không đọc được ID từ cột ẩn → query theo Code
                                var employee = AppDataContext.Employees
                                    .FirstOrDefault(e => e.Code.Equals(code, StringComparison.OrdinalIgnoreCase));

                                if (employee != null)
                                {
                                    employees.Add(new EmployeeImportResult
                                    {
                                        RowNumber = row,
                                        Code = code,
                                        EmployeeId = employee.Id,
                                        EmployeeCvId = employee.EmployeeId,
                                        FullName = AppDataContext.EmployeeCvs
                                            .FirstOrDefault(cv => cv.Id == employee.EmployeeId)?.FullName
                                    });
                                }
                                else
                                {
                                    errors.Add($"Dòng {row}: Mã '{code}' không tồn tại");
                                }
                            }
                        }

                        row++;
                    }
                }
            }

            return Ok(new
            {
                message = errors.Any() ? $"⚠️ Có {errors.Count} lỗi" : "✅ Import thành công",
                success = !errors.Any(),
                totalRows = employees.Count + errors.Count,
                validCount = employees.Count,
                invalidCount = errors.Count,
                employees = employees,
                errors = errors
            });
        }

        public class EmployeeImportResult
        {
            public int RowNumber { get; set; }
            public string Code { get; set; }
            public long EmployeeId { get; set; }
            public long EmployeeCvId { get; set; }
            public string FullName { get; set; }
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