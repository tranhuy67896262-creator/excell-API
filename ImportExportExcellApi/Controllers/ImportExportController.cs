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

                // Header cho sheet lookup
                lookupWs.Cells[1, 1].Value = "CODE";
                lookupWs.Cells[1, 2].Value = "FULL_NAME";
                lookupWs.Cells[1, 3].Value = "EMPLOYEE_ID";
                lookupWs.Cells[1, 4].Value = "CERT_NAME";
                lookupWs.Cells[1, 5].Value = "CERT_ID";
                lookupWs.Cells[1, 6].Value = "YES_NO";
                lookupWs.Cells[1, 7].Value = "GRAD_SCHOOL_NAME";
                lookupWs.Cells[1, 8].Value = "GRAD_SCHOOL_ID";
                lookupWs.Cells[1, 1, 1, 8].Style.Font.Bold = true;

                // Đổ dữ liệu Employee
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

                // Đổ dữ liệu Certificate Type
                var certificateTypes = AppDataContext.SysOtherLists
                    .Where(s => s.TypeCode == "CERTIFICATE_TYPE")
                    .OrderBy(s => s.Name)
                    .ToList();

                for (int i = 0; i < certificateTypes.Count; i++)
                {
                    lookupWs.Cells[i + 2, 4].Value = certificateTypes[i].Name;
                    lookupWs.Cells[i + 2, 5].Value = certificateTypes[i].Id;
                }

                // Tạo list Có/Không cho dropdown
                lookupWs.Cells[2, 6].Value = "Có";
                lookupWs.Cells[3, 6].Value = "Không";

                // ✅ Đổ dữ liệu Đơn vị đào tạo (GRADUATE_SCHOOL)
                var graduateSchools = AppDataContext.SysOtherLists
                    .Where(s => s.TypeCode == "GRADUATE_SCHOOL")
                    .OrderBy(s => s.Name)
                    .ToList();

                for (int i = 0; i < graduateSchools.Count; i++)
                {
                    lookupWs.Cells[i + 2, 7].Value = graduateSchools[i].Name; // Tên hiển thị
                    lookupWs.Cells[i + 2, 8].Value = graduateSchools[i].Id;   // ID ẩn
                }

                // Tạo Named Ranges
                package.Workbook.Names.Add("DanhSachNhanVien", lookupWs.Cells[$"A2:B{lookupData.Count + 1}"]);
                package.Workbook.Names.Add("DanhSachCodeId", lookupWs.Cells[$"A2:C{lookupData.Count + 1}"]);
                package.Workbook.Names.Add("DanhSachCode", lookupWs.Cells[$"A2:A{lookupData.Count + 1}"]);
                package.Workbook.Names.Add("DanhSachChungChi", lookupWs.Cells[$"D2:E{certificateTypes.Count + 1}"]);
                package.Workbook.Names.Add("DanhSachChungChiName", lookupWs.Cells[$"D2:D{certificateTypes.Count + 1}"]);
                package.Workbook.Names.Add("DanhSachCoKhong", lookupWs.Cells[$"F2:F3"]);

                // ✅ Named Range cho Đơn vị đào tạo
                package.Workbook.Names.Add("DanhSachDonViDaoTao", lookupWs.Cells[$"G2:H{graduateSchools.Count + 1}"]);
                package.Workbook.Names.Add("DanhSachDonViDaoTaoName", lookupWs.Cells[$"G2:G{graduateSchools.Count + 1}"]);

                // ========================================
                // 2. XỬ LÝ SHEET CHÍNH
                // ========================================
                var ws = package.Workbook.Worksheets[0];
                int startRow = 6;
                int endRow = 100;

                // ✅ Dropdown Cột A (1): Mã nhân viên
                ApplyDropdown(ws, startRow, endRow, 1, "DanhSachCode");

                // ✅ Dropdown Cột C (3): Loại bằng cấp/Chứng chỉ
                ApplyDropdown(ws, startRow, endRow, 3, "DanhSachChungChiName");

                // ✅ Dropdown Cột D (4): Là bằng chính? ["Có", "Không"]
                ApplyDropdown(ws, startRow, endRow, 4, "DanhSachCoKhong");

                // ✅ Cột E (5): Tên bằng cấp/Chứng chỉ - TEXT INPUT (không dropdown)
                // User nhập tay

                // ✅ Dropdown Cột F (6): Đơn vị đào tạo 🎯
                ApplyDropdown(ws, startRow, endRow, 6, "DanhSachDonViDaoTaoName");

                // ✅ Công thức cho các cột
                for (int row = startRow; row <= endRow; row++)
                {
                    // Cột B (2): FullName - VLOOKUP từ Code
                    string nameFormula = $"=IF(A{row}=\"\", \"\", VLOOKUP(A{row}, DanhSachNhanVien, 2, FALSE))";
                    ws.Cells[row, 2].Formula = nameFormula;

                    // Cột Z (26): Hidden EmployeeId
                    string empIdFormula = $"=IF(A{row}=\"\", \"\", VLOOKUP(A{row}, DanhSachCodeId, 3, FALSE))";
                    ws.Cells[row, 26].Formula = empIdFormula;

                    // Cột AA (27): Hidden CertificateTypeId
                    string certIdFormula = $"=IF(C{row}=\"\", \"\", VLOOKUP(C{row}, DanhSachChungChi, 2, FALSE))";
                    ws.Cells[row, 27].Formula = certIdFormula;

                    // ✅ Cột AB (28): Hidden GraduateSchoolId - VLOOKUP từ tên trường (cột F)
                    string schoolIdFormula = $"=IF(F{row}=\"\", \"\", VLOOKUP(F{row}, DanhSachDonViDaoTao, 2, FALSE))";
                    ws.Cells[row, 28].Formula = schoolIdFormula;
                }

                // 🔐 Ẩn các cột chứa ID
                ws.Column(26).Hidden = true; // EmployeeId
                ws.Column(26).Width = 0.1;
                ws.Column(27).Hidden = true; // CertificateTypeId
                ws.Column(27).Width = 0.1;
                ws.Column(28).Hidden = true; // GraduateSchoolId
                ws.Column(28).Width = 0.1;

                // Border cho vùng dữ liệu
                var dataRange = ws.Cells[$"A{startRow}:Y{endRow}"];
                dataRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                dataRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                dataRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                dataRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;

                // Format header
                ws.Cells[4, 4].Value = "Là bằng chính";
                ws.Cells[4, 4].Style.Font.Bold = true;
                ws.Cells[4, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                ws.Cells[4, 5].Value = "Tên bằng cấp/Chứng chỉ"; // Header cột E - TEXT
                ws.Cells[4, 5].Style.Font.Bold = true;
                ws.Cells[4, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                ws.Cells[4, 6].Value = "Đơn vị đào tạo"; // Header cột F - DROPDOWN
                ws.Cells[4, 6].Style.Font.Bold = true;
                ws.Cells[4, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

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
                    int codeColumn = 1;           // Cột A: Code
                    int certNameColumn = 3;       // Cột C: Certificate Name
                    int isPrimaryColumn = 4;      // Cột D: Là bằng chính
                    int certTextInputColumn = 5;  // ✅ Cột E: Tên bằng cấp (Text input)
                    int schoolNameColumn = 6;     // ✅ Cột F: Đơn vị đào tạo (Dropdown)
                    int hiddenEmpIdColumn = 26;   // Cột Z: EmployeeId
                    int hiddenCertIdColumn = 27;  // Cột AA: CertificateTypeId
                    int hiddenSchoolIdColumn = 28;// Cột AB: GraduateSchoolId

                    while (ws.Cells[row, codeColumn].Value != null)
                    {
                        var code = ws.Cells[row, codeColumn].Value?.ToString()?.Trim();

                        if (!string.IsNullOrEmpty(code))
                        {
                            // Validate Employee
                            var employee = AppDataContext.Employees
                                .FirstOrDefault(e => e.Code.Equals(code, StringComparison.OrdinalIgnoreCase));

                            if (employee == null)
                            {
                                errors.Add($"Dòng {row}: Mã '{code}' không tồn tại");
                                row++;
                                continue;
                            }

                            // Đọc CertificateTypeId
                            long? certificateTypeId = null;
                            var certName = ws.Cells[row, certNameColumn].Value?.ToString()?.Trim();
                            var certIdValue = ws.Cells[row, hiddenCertIdColumn].Value;

                            if (!string.IsNullOrEmpty(certName) && certIdValue != null && long.TryParse(certIdValue.ToString(), out long parsedCertId))
                            {
                                var cert = AppDataContext.SysOtherLists.FirstOrDefault(c => c.Id == parsedCertId && c.Name == certName);
                                certificateTypeId = cert?.Id;
                            }

                            // ✅ Đọc Tên bằng cấp từ cột E (Text input)
                            var certTextName = ws.Cells[row, certTextInputColumn].Value?.ToString()?.Trim();

                            // ✅ Đọc GraduateSchoolId từ cột ẩn AB (từ dropdown cột F)
                            long? graduateSchoolId = null;
                            var schoolName = ws.Cells[row, schoolNameColumn].Value?.ToString()?.Trim();
                            var schoolIdValue = ws.Cells[row, hiddenSchoolIdColumn].Value;

                            if (!string.IsNullOrEmpty(schoolName) && schoolIdValue != null && long.TryParse(schoolIdValue.ToString(), out long parsedSchoolId))
                            {
                                var school = AppDataContext.SysOtherLists
                                    .FirstOrDefault(s => s.Id == parsedSchoolId && s.Name == schoolName && s.TypeCode == "GRADUATE_SCHOOL");
                                graduateSchoolId = school?.Id;
                            }

                            // Đọc giá trị "Là bằng chính"
                            var isPrimaryRaw = ws.Cells[row, isPrimaryColumn].Value?.ToString()?.Trim();
                            bool isPrimary = isPrimaryRaw?.Equals("Có", StringComparison.OrdinalIgnoreCase) == true;

                            employees.Add(new EmployeeImportResult
                            {
                                RowNumber = row,
                                Code = code,
                                EmployeeId = employee.Id,
                                EmployeeCvId = employee.EmployeeId,
                                FullName = AppDataContext.EmployeeCvs.FirstOrDefault(cv => cv.Id == employee.EmployeeId)?.FullName,
                                CertificateName = certName,
                                CertificateTypeId = certificateTypeId,
                                IsPrimaryCertificate = isPrimary,

                                // ✅ Thông tin mới với thứ tự đúng
                                CertificateTextName = certTextName,  // Tên bằng cấp (text input từ cột E)
                                GraduateSchoolName = schoolName,     // Đơn vị đào tạo (dropdown từ cột F)
                                GraduateSchoolId = graduateSchoolId
                            });
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
            public string CertificateName { get; set; }      // Loại bằng cấp (dropdown cột C)
            public long? CertificateTypeId { get; set; }
            public bool IsPrimaryCertificate { get; set; }

            // ✅ Thông tin mới
            public string CertificateTextName { get; set; }  // Tên bằng cấp (text input cột E)
            public string GraduateSchoolName { get; set; }   // Đơn vị đào tạo (dropdown cột F)
            public long? GraduateSchoolId { get; set; }      // ID đơn vị đào tạo
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