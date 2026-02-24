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

        // ========================================
        // ✅ API 1: GEN RELEASE FILE (Chạy 1 lần)
        // ========================================
        [HttpPost("gen-release-file")]
        public IActionResult GenReleaseFile()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            string templateFolder = Path.Combine(_environment.ContentRootPath, _templateFolderName);
            string baseTemplatePath = Path.Combine(templateFolder, _baseTemplateName);
            string releaseTemplatePath = Path.Combine(templateFolder, _releaseTemplateName);

            if (!System.IO.File.Exists(baseTemplatePath))
            {
                return NotFound(new { message = $"Không tìm thấy file Base Template tại: {baseTemplatePath}" });
            }

            using (var package = new ExcelPackage(new FileInfo(baseTemplatePath)))
            {
                // --- BẮT ĐẦU DEEP CLEAN ---
                // Copy sheet cũ sang sheet tạm để dọn sạch lỗi DataValidation cũ
                var oldWs = package.Workbook.Worksheets[0];
                string originalName = oldWs.Name;
                var ws = package.Workbook.Worksheets.Add(originalName + "_Temp", oldWs);
                package.Workbook.Worksheets.Delete(oldWs);
                ws.Name = originalName;
                // --- KẾT THÚC DEEP CLEAN ---

                string lookupSheetName = "Data_Lookup";

                // Xóa sheet lookup cũ nếu có
                if (package.Workbook.Worksheets.Any(w => w.Name == lookupSheetName))
                {
                    package.Workbook.Worksheets.Delete(lookupSheetName);
                }

                var lookupWs = package.Workbook.Worksheets.Add(lookupSheetName);
                lookupWs.Hidden = eWorkSheetHidden.Hidden;

                // === 1. TẠO HEADER CHO LOOKUP SHEET ===
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

                // Static data: Yes/No
                lookupWs.Cells[2, 6].Value = "Có";
                lookupWs.Cells[3, 6].Value = "Không";

                // === 2. TẠO NAMED RANGES VỚI VÙNG CỐ ĐỊNH LỚN (A2:A1000) ===
                // ✅ Dùng vùng lớn để chứa tối đa 1000 items, không cần update lại khi export
                AddOrReplaceNamedRange(package.Workbook, "DanhSachNhanVien", lookupWs.Cells["A2:B1000"]);
                AddOrReplaceNamedRange(package.Workbook, "DanhSachCodeId", lookupWs.Cells["A2:C1000"]);
                AddOrReplaceNamedRange(package.Workbook, "DanhSachCode", lookupWs.Cells["A2:A1000"]);
                AddOrReplaceNamedRange(package.Workbook, "DanhSachChungChi", lookupWs.Cells["D2:E1000"]);
                AddOrReplaceNamedRange(package.Workbook, "DanhSachChungChiName", lookupWs.Cells["D2:D1000"]);
                AddOrReplaceNamedRange(package.Workbook, "DanhSachCoKhong", lookupWs.Cells["F2:F3"]);
                AddOrReplaceNamedRange(package.Workbook, "DanhSachDonViDaoTao", lookupWs.Cells["G2:H1000"]);
                AddOrReplaceNamedRange(package.Workbook, "DanhSachDonViDaoTaoName", lookupWs.Cells["G2:G1000"]);
                AddOrReplaceNamedRange(package.Workbook, "DanhSachLevelId", lookupWs.Cells["I2:J1000"]);
                AddOrReplaceNamedRange(package.Workbook, "DanhSachLevelIdName", lookupWs.Cells["I2:I1000"]);
                AddOrReplaceNamedRange(package.Workbook, "DanhSachLevelTrain", lookupWs.Cells["K2:L1000"]);
                AddOrReplaceNamedRange(package.Workbook, "DanhSachLevelTrainName", lookupWs.Cells["K2:K1000"]);
                AddOrReplaceNamedRange(package.Workbook, "DanhSachTrainingMethod", lookupWs.Cells["M2:N1000"]);
                AddOrReplaceNamedRange(package.Workbook, "DanhSachTrainingMethodName", lookupWs.Cells["M2:M1000"]);

                // === 3. ÁP DỤNG DROPDOWN CHO SHEET CHÍNH ===
                int startRow = 6;
                int endRow = 100;

                ApplyDropdown(ws, startRow, endRow, 1, "DanhSachCode");
                ApplyDropdown(ws, startRow, endRow, 3, "DanhSachChungChiName");
                ApplyDropdown(ws, startRow, endRow, 4, "DanhSachCoKhong");
                ApplyDropdown(ws, startRow, endRow, 6, "DanhSachDonViDaoTaoName");
                ApplyDropdown(ws, startRow, endRow, 7, "DanhSachLevelIdName");
                ApplyDropdown(ws, startRow, endRow, 8, "DanhSachLevelTrainName");
                ApplyDropdown(ws, startRow, endRow, 9, "DanhSachTrainingMethodName");

                // === 4. THÊM FORMULA CHO SHEET CHÍNH ===
                for (int row = startRow; row <= endRow; row++)
                {
                    ws.Cells[row, 2].Formula = $"=IF(A{row}=\"\", \"\", VLOOKUP(A{row}, DanhSachNhanVien, 2, FALSE))";
                    ws.Cells[row, 26].Formula = $"=IF(A{row}=\"\", \"\", VLOOKUP(A{row}, DanhSachCodeId, 3, FALSE))";
                    ws.Cells[row, 27].Formula = $"=IF(C{row}=\"\", \"\", VLOOKUP(C{row}, DanhSachChungChi, 2, FALSE))";
                    ws.Cells[row, 28].Formula = $"=IF(F{row}=\"\", \"\", VLOOKUP(F{row}, DanhSachDonViDaoTao, 2, FALSE))";
                    ws.Cells[row, 29].Formula = $"=IF(G{row}=\"\", \"\", VLOOKUP(G{row}, DanhSachLevelId, 2, FALSE))";
                    ws.Cells[row, 30].Formula = $"=IF(H{row}=\"\", \"\", VLOOKUP(H{row}, DanhSachLevelTrain, 2, FALSE))";
                    ws.Cells[row, 31].Formula = $"=IF(I{row}=\"\", \"\", VLOOKUP(I{row}, DanhSachTrainingMethod, 2, FALSE))";
                }

                // === 5. ẨN CỘT ID (Z-AE) ===
                for (int col = 26; col <= 31; col++)
                {
                    ws.Column(col).Hidden = true;
                    ws.Column(col).Width = 0.1;
                }

                // === 6. FORMAT SỐ/NGÀY ===
                ws.Column(10).Style.Numberformat.Format = "0";
                ws.Column(12).Style.Numberformat.Format = "0.0";
                for (int col = 13; col <= 16; col++)
                    ws.Column(col).Style.Numberformat.Format = "dd/mm/yyyy";

                // === 7. BORDER ===
                var dataRange = ws.Cells[$"A{startRow}:R{endRow}"];
                dataRange.Style.Border.Top.Style = dataRange.Style.Border.Bottom.Style =
                dataRange.Style.Border.Left.Style = dataRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;

                // === 8. LƯU FILE RELEASE (Không ghi đè file Base) ===
                package.SaveAs(new FileInfo(releaseTemplatePath));
            }

            return Ok(new { message = "✅ Release Template đã sẵn sàng", path = releaseTemplatePath });
        }

        // ========================================
        // ✅ API 2: GEN EXCEL TEMPLATE (Production)
        // ========================================
        [HttpGet("gen-excel-template")]
        public IActionResult GenExcelTemplate()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            string templateFolder = Path.Combine(_environment.ContentRootPath, _templateFolderName);
            string releaseTemplatePath = Path.Combine(templateFolder, _releaseTemplateName);

            if (!System.IO.File.Exists(releaseTemplatePath))
            {
                return BadRequest(new { message = "Chưa có file Release Template. Hãy chạy gen-release-file trước." });
            }

            using (var package = new ExcelPackage(new FileInfo(releaseTemplatePath)))
            {
                var lookupWs = package.Workbook.Worksheets.FirstOrDefault(w => w.Name == "Data_Lookup");
                if (lookupWs == null)
                    return BadRequest(new { message = "Base Template thiếu sheet Data_Lookup" });

                lookupWs.Hidden = eWorkSheetHidden.Hidden;

                // === ✅ CLEAR DATA CŨ TRONG LOOKUP SHEET ===
                // Giữ lại header row 1, clear từ row 2 đến 1000
                lookupWs.Cells["A2:N1000"].Clear();

                // === ✅ ĐỔ DATA MỚI TỪ DATABASE ===

                // Employee
                var employees = AppDataContext.Employees
                    .Join(AppDataContext.EmployeeCvs, e => e.EmployeeId, c => c.Id, (e, c) => new { e.Code, e.Id, c.FullName })
                    .OrderBy(x => x.Code).ToList();
                for (int i = 0; i < employees.Count; i++)
                {
                    lookupWs.Cells[i + 2, 1].Value = employees[i].Code;
                    lookupWs.Cells[i + 2, 2].Value = employees[i].FullName;
                    lookupWs.Cells[i + 2, 3].Value = employees[i].Id;
                }

                // Certificate
                var certTypes = AppDataContext.SysOtherLists.Where(s => s.TypeCode == "CERTIFICATE_TYPE").OrderBy(s => s.Name).ToList();
                for (int i = 0; i < certTypes.Count; i++)
                {
                    lookupWs.Cells[i + 2, 4].Value = certTypes[i].Name;
                    lookupWs.Cells[i + 2, 5].Value = certTypes[i].Id;
                }

                // Yes/No (static)
                lookupWs.Cells[2, 6].Value = "Có";
                lookupWs.Cells[3, 6].Value = "Không";

                // Graduate School
                var gradSchools = AppDataContext.SysOtherLists.Where(s => s.TypeCode == "GRADUATE_SCHOOL").OrderBy(s => s.Name).ToList();
                for (int i = 0; i < gradSchools.Count; i++)
                {
                    lookupWs.Cells[i + 2, 7].Value = gradSchools[i].Name;
                    lookupWs.Cells[i + 2, 8].Value = gradSchools[i].Id;
                }

                // Level ID
                var levelIds = AppDataContext.SysOtherLists.Where(s => s.TypeCode == "LEVEL_ID").OrderBy(s => s.Name).ToList();
                for (int i = 0; i < levelIds.Count; i++)
                {
                    lookupWs.Cells[i + 2, 9].Value = levelIds[i].Name;
                    lookupWs.Cells[i + 2, 10].Value = levelIds[i].Id;
                }

                // Level Train
                var levelTrains = AppDataContext.SysOtherLists.Where(s => s.TypeCode == "LEVEL_TRAIN").OrderBy(s => s.Name).ToList();
                for (int i = 0; i < levelTrains.Count; i++)
                {
                    lookupWs.Cells[i + 2, 11].Value = levelTrains[i].Name;
                    lookupWs.Cells[i + 2, 12].Value = levelTrains[i].Id;
                }

                // Training Method
                var trainMethods = AppDataContext.SysOtherLists.Where(s => s.TypeCode == "TRAINING_METHOD").OrderBy(s => s.Name).ToList();
                for (int i = 0; i < trainMethods.Count; i++)
                {
                    lookupWs.Cells[i + 2, 13].Value = trainMethods[i].Name;
                    lookupWs.Cells[i + 2, 14].Value = trainMethods[i].Id;
                }

                // === ✅ (Tuỳ chọn) Điền sample data vào sheet chính ===
                //var ws = package.Workbook.Worksheets[0];
                //for (int i = 0; i < Math.Min(employees.Count, 20); i++)
                //{
                //    ws.Cells[6 + i, 1].Value = employees[i].Code;
                //}

                return File(package.GetAsByteArray(),
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    $"Mau_Nhap_Lieu_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
            }
        }

        // ========================================
        // ✅ API 3: IMPORT EMPLOYEE (Production)
        // ========================================
        [HttpPost("import-employee")]
        public IActionResult ImportEmployee(IFormFile file)
        {
            if (file == null || file.Length == 0)
                return BadRequest(new { message = "Vui lòng chọn file Excel để import." });

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var results = new List<EmployeeImportResult>();
            var errors = new List<string>();

            using (var stream = new MemoryStream())
            {
                file.CopyTo(stream);
                using (var package = new ExcelPackage(stream))
                {
                    var ws = package.Workbook.Worksheets[0];
                    int row = 6;

                    while (ws.Cells[row, 1].Value != null)
                    {
                        var code = ws.Cells[row, 1].Value?.ToString()?.Trim();
                        if (!string.IsNullOrEmpty(code))
                        {
                            var emp = AppDataContext.Employees.FirstOrDefault(e => e.Code.Equals(code, StringComparison.OrdinalIgnoreCase));
                            if (emp == null) { errors.Add($"Dòng {row}: Mã '{code}' không tồn tại"); row++; continue; }

                            results.Add(new EmployeeImportResult
                            {
                                RowNumber = row,
                                Code = code,
                                EmployeeId = emp.Id,
                                EmployeeCvId = emp.EmployeeId,
                                FullName = AppDataContext.EmployeeCvs.FirstOrDefault(c => c.Id == emp.EmployeeId)?.FullName,
                                CertificateName = ws.Cells[row, 3].Value?.ToString()?.Trim(),
                                IsPrimaryCertificate = ws.Cells[row, 4].Value?.ToString()?.Trim() == "Có",
                                CertificateTextName = ws.Cells[row, 5].Value?.ToString()?.Trim(),
                                GraduateSchoolName = ws.Cells[row, 6].Value?.ToString()?.Trim(),
                                LevelIdName = ws.Cells[row, 7].Value?.ToString()?.Trim(),
                                LevelTrainName = ws.Cells[row, 8].Value?.ToString()?.Trim(),
                                TrainingMethodName = ws.Cells[row, 9].Value?.ToString()?.Trim(),
                                Year = ReadInt(ws.Cells[row, 10].Value),
                                ContentTrain = ws.Cells[row, 11].Value?.ToString()?.Trim(),
                                Mark = ReadDecimal(ws.Cells[row, 12].Value),
                                TrainFromDate = ReadDate(ws.Cells[row, 13].Value),
                                TrainToDate = ReadDate(ws.Cells[row, 14].Value),
                                EffectFromDate = ReadDate(ws.Cells[row, 15].Value),
                                EffectToDate = ReadDate(ws.Cells[row, 16].Value),
                                Classification = ws.Cells[row, 17].Value?.ToString()?.Trim(),
                                Remark = ws.Cells[row, 18].Value?.ToString()?.Trim(),
                                CertificateTypeId = ReadHiddenId(ws.Cells[row, 27].Value, "CERTIFICATE_TYPE"),
                                GraduateSchoolId = ReadHiddenId(ws.Cells[row, 28].Value, "GRADUATE_SCHOOL"),
                                LevelId = ReadHiddenId(ws.Cells[row, 29].Value, "LEVEL_ID"),
                                LevelTrainId = ReadHiddenId(ws.Cells[row, 30].Value, "LEVEL_TRAIN"),
                                TrainingMethodId = ReadHiddenId(ws.Cells[row, 31].Value, "TRAINING_METHOD")
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
                totalRows = results.Count + errors.Count,
                validCount = results.Count,
                employees = results,
                errors = errors
            });
        }

        // ========================================
        // ✅ HELPER METHODS
        // ========================================

        /// <summary>
        /// Thêm hoặc thay thế Named Range với ExcelRangeBase
        /// </summary>
        private void AddOrReplaceNamedRange(ExcelWorkbook workbook, string name, ExcelRangeBase range)
        {
            if (workbook.Names.Any(n => n.Name.Equals(name, StringComparison.OrdinalIgnoreCase)))
            {
                workbook.Names.Remove(name); // ✅ Remove bằng string name
            }
            workbook.Names.Add(name, range);
        }

        /// <summary>
        /// Áp dụng Data Validation dropdown
        /// </summary>
        private void ApplyDropdown(ExcelWorksheet ws, int startRow, int endRow, int column, string namedRange)
        {
            if (!ws.Workbook.Names.Any(n => n.Name.Equals(namedRange, StringComparison.OrdinalIgnoreCase)))
                return;

            var range = ws.Cells[startRow, column, endRow, column];

            // Xóa validation cũ nếu có (tránh trùng)
            var existingValidations = ws.DataValidations
                .Where(v => v.Address.Start.Column == column &&
                           v.Address.Start.Row >= startRow &&
                           v.Address.End.Row <= endRow)
                .ToList();
            foreach (var v in existingValidations)
            {
                ws.DataValidations.Remove(v);
            }

            var validation = range.DataValidation.AddListDataValidation();
            validation.Formula.ExcelFormula = namedRange; // ✅ EPPlus tự thêm dấu = khi cần
            validation.ShowErrorMessage = true;
            validation.ErrorStyle = ExcelDataValidationWarningStyle.warning;
            validation.ErrorTitle = "Giá trị không hợp lệ";
            validation.Error = "Vui lòng chọn giá trị có trong danh sách.";
            validation.ShowInputMessage = true;
            validation.PromptTitle = "💡 Hướng dẫn";
            validation.Prompt = "Chọn từ dropdown hoặc paste giá trị hợp lệ.";
        }

        private int? ReadInt(object val) => val != null && int.TryParse(val.ToString(), out int r) ? r : null;

        private decimal? ReadDecimal(object val) => val != null && decimal.TryParse(val.ToString()?.Replace(',', '.'),
            System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out decimal r) ? r : null;

        private DateTime? ReadDate(object val)
        {
            if (val == null) return null;
            if (val is DateTime dt) return dt;
            if (val is double d) { try { return DateTime.FromOADate(d); } catch { } }
            if (val is string s && DateTime.TryParse(s, out DateTime r)) return r;
            return null;
        }

        private long? ReadHiddenId(object val, string typeCode)
        {
            if (val == null || !long.TryParse(val.ToString(), out long id)) return null;
            return AppDataContext.SysOtherLists.FirstOrDefault(s => s.Id == id && s.TypeCode == typeCode)?.Id;
        }
    }

    // ========================================
    // ✅ DTO
    // ========================================
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
        public string Classification { get; set; }
        public string Remark { get; set; }
    }
}