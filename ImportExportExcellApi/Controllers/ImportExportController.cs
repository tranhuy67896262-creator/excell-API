using ImportExportExcellApi.Data;
using ImportExportExcellApi.Entities;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.Style;
using System;
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
                lookupWs.Cells[1, 7].Value = "GRAD_NAME";
                lookupWs.Cells[1, 8].Value = "GRAD_ID";
                lookupWs.Cells[1, 9].Value = "LEVEL_ID_NAME";
                lookupWs.Cells[1, 10].Value = "LEVEL_ID_VAL";
                lookupWs.Cells[1, 11].Value = "LEVEL_TRAIN_NAME";
                lookupWs.Cells[1, 12].Value = "LEVEL_TRAIN_VAL";
                lookupWs.Cells[1, 13].Value = "TRAIN_METHOD_NAME";
                lookupWs.Cells[1, 14].Value = "TRAIN_METHOD_VAL";
                lookupWs.Cells[1, 1, 1, 14].Style.Font.Bold = true;

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

                // Certificate Type
                var certificateTypes = AppDataContext.SysOtherLists
                    .Where(s => s.TypeCode == "CERTIFICATE_TYPE")
                    .OrderBy(s => s.Name)
                    .ToList();
                for (int i = 0; i < certificateTypes.Count; i++)
                {
                    lookupWs.Cells[i + 2, 4].Value = certificateTypes[i].Name;
                    lookupWs.Cells[i + 2, 5].Value = certificateTypes[i].Id;
                }

                // Có/Không
                lookupWs.Cells[2, 6].Value = "Có";
                lookupWs.Cells[3, 6].Value = "Không";

                // Graduate School
                var graduateSchools = AppDataContext.SysOtherLists
                    .Where(s => s.TypeCode == "GRADUATE_SCHOOL")
                    .OrderBy(s => s.Name)
                    .ToList();
                for (int i = 0; i < graduateSchools.Count; i++)
                {
                    lookupWs.Cells[i + 2, 7].Value = graduateSchools[i].Name;
                    lookupWs.Cells[i + 2, 8].Value = graduateSchools[i].Id;
                }

                // LEVEL_ID (Trình độ chuyên môn)
                var levelIds = AppDataContext.SysOtherLists
                    .Where(s => s.TypeCode == "LEVEL_ID")
                    .OrderBy(s => s.Name)
                    .ToList();
                for (int i = 0; i < levelIds.Count; i++)
                {
                    lookupWs.Cells[i + 2, 9].Value = levelIds[i].Name;
                    lookupWs.Cells[i + 2, 10].Value = levelIds[i].Id;
                }

                // LEVEL_TRAIN (Trình độ học vấn)
                var levelTrains = AppDataContext.SysOtherLists
                    .Where(s => s.TypeCode == "LEVEL_TRAIN")
                    .OrderBy(s => s.Name)
                    .ToList();
                for (int i = 0; i < levelTrains.Count; i++)
                {
                    lookupWs.Cells[i + 2, 11].Value = levelTrains[i].Name;
                    lookupWs.Cells[i + 2, 12].Value = levelTrains[i].Id;
                }

                // TRAINING_METHOD (Hình thức đào tạo)
                var trainingMethods = AppDataContext.SysOtherLists
                    .Where(s => s.TypeCode == "TRAINING_METHOD")
                    .OrderBy(s => s.Name)
                    .ToList();
                for (int i = 0; i < trainingMethods.Count; i++)
                {
                    lookupWs.Cells[i + 2, 13].Value = trainingMethods[i].Name;
                    lookupWs.Cells[i + 2, 14].Value = trainingMethods[i].Id;
                }

                // ✅ XÓA PHẦN CLASSIFICATION - Vì giờ là text input, không cần trong lookup

                // Tạo Named Ranges (XÓA DanhSachXepLoai)
                package.Workbook.Names.Add("DanhSachNhanVien", lookupWs.Cells[$"A2:B{lookupData.Count + 1}"]);
                package.Workbook.Names.Add("DanhSachCodeId", lookupWs.Cells[$"A2:C{lookupData.Count + 1}"]);
                package.Workbook.Names.Add("DanhSachCode", lookupWs.Cells[$"A2:A{lookupData.Count + 1}"]);
                package.Workbook.Names.Add("DanhSachChungChi", lookupWs.Cells[$"D2:E{certificateTypes.Count + 1}"]);
                package.Workbook.Names.Add("DanhSachChungChiName", lookupWs.Cells[$"D2:D{certificateTypes.Count + 1}"]);
                package.Workbook.Names.Add("DanhSachCoKhong", lookupWs.Cells[$"F2:F3"]);
                package.Workbook.Names.Add("DanhSachDonViDaoTao", lookupWs.Cells[$"G2:H{graduateSchools.Count + 1}"]);
                package.Workbook.Names.Add("DanhSachDonViDaoTaoName", lookupWs.Cells[$"G2:G{graduateSchools.Count + 1}"]);
                package.Workbook.Names.Add("DanhSachLevelId", lookupWs.Cells[$"I2:J{levelIds.Count + 1}"]);
                package.Workbook.Names.Add("DanhSachLevelIdName", lookupWs.Cells[$"I2:I{levelIds.Count + 1}"]);
                package.Workbook.Names.Add("DanhSachLevelTrain", lookupWs.Cells[$"K2:L{levelTrains.Count + 1}"]);
                package.Workbook.Names.Add("DanhSachLevelTrainName", lookupWs.Cells[$"K2:K{levelTrains.Count + 1}"]);
                package.Workbook.Names.Add("DanhSachTrainingMethod", lookupWs.Cells[$"M2:N{trainingMethods.Count + 1}"]);
                package.Workbook.Names.Add("DanhSachTrainingMethodName", lookupWs.Cells[$"M2:M{trainingMethods.Count + 1}"]);
                // ✅ XÓA: package.Workbook.Names.Add("DanhSachXepLoai", ...)

                // ========================================
                // 2. XỬ LÝ SHEET CHÍNH - GIỮ NGUYÊN TEMPLATE
                // ========================================
                var ws = package.Workbook.Worksheets[0];
                int startRow = 6;
                int endRow = 100;

                // ✅ CHỈ ÁP DỤNG DROPDOWN CHO CÁC CỘT CẦN THIẾT
                ApplyDropdown(ws, startRow, endRow, 1, "DanhSachCode");              // A: Code
                ApplyDropdown(ws, startRow, endRow, 3, "DanhSachChungChiName");      // C: Loại bằng
                ApplyDropdown(ws, startRow, endRow, 4, "DanhSachCoKhong");           // D: Bằng chính
                // E: Tên bằng (text input)
                ApplyDropdown(ws, startRow, endRow, 6, "DanhSachDonViDaoTaoName");   // F: Đơn vị đào tạo
                ApplyDropdown(ws, startRow, endRow, 7, "DanhSachLevelIdName");       // G: Trình độ chuyên môn
                ApplyDropdown(ws, startRow, endRow, 8, "DanhSachLevelTrainName");    // H: Trình độ học vấn
                ApplyDropdown(ws, startRow, endRow, 9, "DanhSachTrainingMethodName"); // I: Hình thức đào tạo
                // ✅ Q: Xếp loại - TEXT INPUT (không dropdown)
                // ✅ R: Ghi chú - TEXT INPUT (không dropdown)

                // ✅ CHỈ THÊM CÔNG THỨC - KHÔNG SỬA GIÁ TRỊ
                for (int row = startRow; row <= endRow; row++)
                {
                    // FullName (cột B)
                    ws.Cells[row, 2].Formula = $"=IF(A{row}=\"\", \"\", VLOOKUP(A{row}, DanhSachNhanVien, 2, FALSE))";

                    // Hidden EmployeeId (Z=26)
                    ws.Cells[row, 26].Formula = $"=IF(A{row}=\"\", \"\", VLOOKUP(A{row}, DanhSachCodeId, 3, FALSE))";

                    // Hidden CertificateTypeId (AA=27)
                    ws.Cells[row, 27].Formula = $"=IF(C{row}=\"\", \"\", VLOOKUP(C{row}, DanhSachChungChi, 2, FALSE))";

                    // Hidden GraduateSchoolId (AB=28)
                    ws.Cells[row, 28].Formula = $"=IF(F{row}=\"\", \"\", VLOOKUP(F{row}, DanhSachDonViDaoTao, 2, FALSE))";

                    // Hidden LevelId (AC=29)
                    ws.Cells[row, 29].Formula = $"=IF(G{row}=\"\", \"\", VLOOKUP(G{row}, DanhSachLevelId, 2, FALSE))";

                    // Hidden LevelTrain (AD=30)
                    ws.Cells[row, 30].Formula = $"=IF(H{row}=\"\", \"\", VLOOKUP(H{row}, DanhSachLevelTrain, 2, FALSE))";

                    // Hidden TrainingMethod (AE=31)
                    ws.Cells[row, 31].Formula = $"=IF(I{row}=\"\", \"\", VLOOKUP(I{row}, DanhSachTrainingMethod, 2, FALSE))";
                }

                // 🔐 Ẩn các cột chứa ID (26-31)
                for (int col = 26; col <= 31; col++)
                {
                    ws.Column(col).Hidden = true;
                    ws.Column(col).Width = 0.1;
                }

                // ✅ CHỈ FORMAT - KHÔNG SỬA NỘI DUNG
                ws.Column(10).Style.Numberformat.Format = "0";          // YEAR
                ws.Column(12).Style.Numberformat.Format = "0.0";        // MARK
                ws.Column(13).Style.Numberformat.Format = "dd/mm/yyyy"; // TRAIN_FROM_DATE
                ws.Column(14).Style.Numberformat.Format = "dd/mm/yyyy"; // TRAIN_TO_DATE
                ws.Column(15).Style.Numberformat.Format = "dd/mm/yyyy"; // EFFECT_FROM
                ws.Column(16).Style.Numberformat.Format = "dd/mm/yyyy"; // EFFECT_TO

                // Border cho vùng dữ liệu
                var dataRange = ws.Cells[$"A{startRow}:R{endRow}"];
                dataRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                dataRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                dataRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                dataRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;

                // ✅ KHÔNG SỬA HEADER - GIỮ NGUYÊN TỪ TEMPLATE

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

            // ✅ Helper function để đọc ngày tháng từ Excel
            DateTime? ReadDateFromExcel(object cellValue)
            {
                if (cellValue == null) return null;

                // Trường hợp 1: Excel đã lưu là DateTime
                if (cellValue is DateTime dateTime)
                {
                    return dateTime;
                }

                // Trường hợp 2: Excel lưu là số (serial number)
                if (cellValue is double serialNumber)
                {
                    try
                    {
                        return DateTime.FromOADate(serialNumber);
                    }
                    catch
                    {
                        return null;
                    }
                }

                // Trường hợp 3: Excel lưu là string
                if (cellValue is string stringValue)
                {
                    if (DateTime.TryParseExact(stringValue,
                        new[] { "dd/MM/yyyy", "dd-MM-yyyy", "yyyy-MM-dd", "dd/MM/yyyy", "M/d/yyyy" },
                        System.Globalization.CultureInfo.InvariantCulture,
                        System.Globalization.DateTimeStyles.None,
                        out DateTime parsedDate))
                    {
                        return parsedDate;
                    }

                    // Thử parse mặc định
                    if (DateTime.TryParse(stringValue, out parsedDate))
                    {
                        return parsedDate;
                    }
                }

                return null;
            }

            using (var stream = new MemoryStream())
            {
                file.CopyTo(stream);
                using (var package = new ExcelPackage(stream))
                {
                    var ws = package.Workbook.Worksheets[0];
                    int startRow = 6;
                    int row = startRow;

                    while (ws.Cells[row, 1].Value != null)
                    {
                        var code = ws.Cells[row, 1].Value?.ToString()?.Trim();

                        if (!string.IsNullOrEmpty(code))
                        {
                            var employee = AppDataContext.Employees
                                .FirstOrDefault(e => e.Code.Equals(code, StringComparison.OrdinalIgnoreCase));

                            if (employee == null)
                            {
                                errors.Add($"Dòng {row}: Mã '{code}' không tồn tại");
                                row++;
                                continue;
                            }

                            // Helper đọc ID từ cột ẩn
                            long? ReadHiddenId(int nameCol, int hiddenCol, string typeCode)
                            {
                                var name = ws.Cells[row, nameCol].Value?.ToString()?.Trim();
                                var idVal = ws.Cells[row, hiddenCol].Value;
                                if (!string.IsNullOrEmpty(name) && idVal != null && long.TryParse(idVal.ToString(), out long parsedId))
                                {
                                    var item = AppDataContext.SysOtherLists
                                        .FirstOrDefault(s => s.Id == parsedId && s.Name == name && s.TypeCode == typeCode);
                                    return item?.Id;
                                }
                                return null;
                            }

                            // Đọc các trường dữ liệu
                            int? year = null;
                            if (ws.Cells[row, 10].Value != null && int.TryParse(ws.Cells[row, 10].Value.ToString(), out int parsedYear))
                            {
                                year = parsedYear;
                            }

                            decimal? mark = null;
                            if (ws.Cells[row, 12].Value != null && decimal.TryParse(ws.Cells[row, 12].Value.ToString().Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out decimal parsedMark))
                            {
                                mark = parsedMark;
                            }

                            // ✅ Đọc ngày tháng với helper mới
                            DateTime? trainFromDate = ReadDateFromExcel(ws.Cells[row, 13].Value);
                            DateTime? trainToDate = ReadDateFromExcel(ws.Cells[row, 14].Value);
                            DateTime? effectFromDate = ReadDateFromExcel(ws.Cells[row, 15].Value);
                            DateTime? effectToDate = ReadDateFromExcel(ws.Cells[row, 16].Value);

                            employees.Add(new EmployeeImportResult
                            {
                                RowNumber = row,
                                Code = code,
                                EmployeeId = employee.Id,
                                EmployeeCvId = employee.EmployeeId,
                                FullName = AppDataContext.EmployeeCvs.FirstOrDefault(cv => cv.Id == employee.EmployeeId)?.FullName,
                                CertificateName = ws.Cells[row, 3].Value?.ToString()?.Trim(),
                                CertificateTypeId = ReadHiddenId(3, 27, "CERTIFICATE_TYPE"),
                                IsPrimaryCertificate = ws.Cells[row, 4].Value?.ToString()?.Trim()?.Equals("Có", StringComparison.OrdinalIgnoreCase) == true,
                                CertificateTextName = ws.Cells[row, 5].Value?.ToString()?.Trim(),
                                GraduateSchoolName = ws.Cells[row, 6].Value?.ToString()?.Trim(),
                                GraduateSchoolId = ReadHiddenId(6, 28, "GRADUATE_SCHOOL"),
                                LevelIdName = ws.Cells[row, 7].Value?.ToString()?.Trim(),
                                LevelId = ReadHiddenId(7, 29, "LEVEL_ID"),
                                LevelTrainName = ws.Cells[row, 8].Value?.ToString()?.Trim(),
                                LevelTrainId = ReadHiddenId(8, 30, "LEVEL_TRAIN"),
                                TrainingMethodName = ws.Cells[row, 9].Value?.ToString()?.Trim(),
                                TrainingMethodId = ReadHiddenId(9, 31, "TRAINING_METHOD"),
                                Year = year,
                                ContentTrain = ws.Cells[row, 11].Value?.ToString()?.Trim(),
                                Mark = mark,
                                TrainFromDate = trainFromDate,
                                TrainToDate = trainToDate,
                                EffectFromDate = effectFromDate,
                                EffectToDate = effectToDate,
                                Classification = ws.Cells[row, 17].Value?.ToString()?.Trim(),
                                Remark = ws.Cells[row, 18].Value?.ToString()?.Trim()
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
            public string CertificateName { get; set; }
            public long? CertificateTypeId { get; set; }
            public bool IsPrimaryCertificate { get; set; }
            public string CertificateTextName { get; set; }
            public string GraduateSchoolName { get; set; }
            public long? GraduateSchoolId { get; set; }
            public string LevelIdName { get; set; }
            public long? LevelId { get; set; }
            public string LevelTrainName { get; set; }
            public long? LevelTrainId { get; set; }
            public string TrainingMethodName { get; set; }
            public long? TrainingMethodId { get; set; }
            public int? Year { get; set; }
            public string ContentTrain { get; set; }
            public decimal? Mark { get; set; }
            public DateTime? TrainFromDate { get; set; }
            public DateTime? TrainToDate { get; set; }
            public DateTime? EffectFromDate { get; set; }
            public DateTime? EffectToDate { get; set; }
            public string Classification { get; set; }  // ✅ TEXT INPUT
            public string Remark { get; set; }           // ✅ TEXT INPUT
        }

        private void ApplyDropdown(ExcelWorksheet ws, int startRow, int endRow, int column, string namedRange)
        {
            if (ws.Workbook.Names.Any(n => n.Name.Equals(namedRange, StringComparison.OrdinalIgnoreCase)))
            {
                var range = ws.Cells[startRow, column, endRow, column];
                var validation = range.DataValidation.AddListDataValidation();
                validation.Formula.ExcelFormula = $"INDIRECT(\"{namedRange}\")";
                validation.ShowErrorMessage = true;
                validation.ErrorStyle = ExcelDataValidationWarningStyle.warning;
                validation.ErrorTitle = "Giá trị không hợp lệ";
                validation.Error = "Vui lòng chọn giá trị có trong danh sách.";
                validation.ShowInputMessage = true;
                validation.PromptTitle = "💡 Hướng dẫn";
                validation.Prompt = "Chọn từ dropdown hoặc paste giá trị hợp lệ.";
            }
        }
    }
}