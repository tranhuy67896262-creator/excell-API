using System;
using System.Collections.Generic; // Thêm cái này
using System.Drawing;
using System.IO;
using System.Linq;
using ImportExportExcellApi.Data;
using ImportExportExcellApi.Entities;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.Style;

namespace ImportExportExcellApi.Controllers;

[Route("api/[controller]")]
[ApiController]
public class HuWorkingController : ControllerBase
{
    private readonly string _baseTemplateName = "HU_WORKING_BASE.xlsx";
    private readonly IWebHostEnvironment _environment;
    private readonly string _releaseTemplateName = "HU_WORKING_BASE_Release.xlsx";
    private readonly string _templateFolderName = "Templates";

    public HuWorkingController(IWebHostEnvironment environment)
    {
        _environment = environment;
        AppDataContext.Initialize();

        string templatePath = Path.Combine(_environment.ContentRootPath, _templateFolderName);
        if (!Directory.Exists(templatePath))
        {
            Directory.CreateDirectory(templatePath);
        }
    }

    //  Đã xóa các property Code, FullName, Id thừa ở đây
    [HttpGet("export")]
    public IActionResult Export()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        var path = Path.Combine(_environment.ContentRootPath, _templateFolderName, _baseTemplateName);

        using (var package = new ExcelPackage(new FileInfo(path)))
        {
            var ws = package.Workbook.Worksheets[0];
            var lookup = package.Workbook.Worksheets.FirstOrDefault(w => w.Name == "Data_Lookup") ??
                         package.Workbook.Worksheets.Add("Data_Lookup");

            lookup.Hidden = eWorkSheetHidden.VeryHidden;
            lookup.Cells.Clear();

            var emps = AppDataContext.Employees
                .Join(AppDataContext.EmployeeCvs, e => e.EmployeeId, c => c.Id,
                    (e, c) => new { e.Code, c.FullName, e.Id })
                .OrderBy(x => x.Code)
                .ToList();

            for (int i = 0; i < emps.Count; i++)
            {
                lookup.Cells[i + 2, 1].Value = emps[i].Code;
                lookup.Cells[i + 2, 2].Value = emps[i].FullName;
                lookup.Cells[i + 2, 3].Value = emps[i].Id;
            }

            // TYPE_DECISION (Loại Quyết Định) -> Name (D), Id (E)
            var decisions = AppDataContext.SysOtherLists
                .Where(s => s.TypeCode == "TYPE_DECISION")
                .OrderBy(s => s.Name)
                .ToList();
            for (int i = 0; i < decisions.Count; i++)
            {
                lookup.Cells[i + 2, 4].Value = decisions[i].Name;
                lookup.Cells[i + 2, 5].Value = decisions[i].Id;
            }

            // Phòng ban (F, G)
            var depts = AppDataContext.Departments.OrderBy(d => d.Name).ToList();
            for (int i = 0; i < depts.Count; i++)
            {
                lookup.Cells[i + 2, 6].Value = depts[i].Name;
                lookup.Cells[i + 2, 7].Value = depts[i].Id;
            }

            // Chức danh (H, I)
            var positions = AppDataContext.Positions.OrderBy(p => p.Name).ToList();
            for (int i = 0; i < positions.Count; i++)
            {
                lookup.Cells[i + 2, 8].Value = positions[i].Name;
                lookup.Cells[i + 2, 9].Value = positions[i].Id;
            }

            var wb = package.Workbook;
            // --- Employee ---
            if (wb.Names.Any(n => n.Name.Equals("DanhSachCode", StringComparison.OrdinalIgnoreCase))) wb.Names.Remove("DanhSachCode");
            if (emps.Any()) wb.Names.Add("DanhSachCode", lookup.Cells[2, 1, 1 + emps.Count, 1]);

            if (wb.Names.Any(n => n.Name.Equals("DanhSachNhanVien", StringComparison.OrdinalIgnoreCase))) wb.Names.Remove("DanhSachNhanVien");
            wb.Names.Add("DanhSachNhanVien", lookup.Cells["A2:B1000"]);

            if (wb.Names.Any(n => n.Name.Equals("DanhSachCodeId", StringComparison.OrdinalIgnoreCase))) wb.Names.Remove("DanhSachCodeId");
            wb.Names.Add("DanhSachCodeId", lookup.Cells["A2:C1000"]);

            // --- TYPE_DECISION ---
            if (wb.Names.Any(n => n.Name.Equals("DanhSachLoaiQuyetDinh", StringComparison.OrdinalIgnoreCase))) wb.Names.Remove("DanhSachLoaiQuyetDinh");
            if (decisions.Any()) wb.Names.Add("DanhSachLoaiQuyetDinh", lookup.Cells[2, 4, 1 + decisions.Count, 4]);
            
            if (wb.Names.Any(n => n.Name.Equals("DanhSachLoaiQuyetDinhRange", StringComparison.OrdinalIgnoreCase))) wb.Names.Remove("DanhSachLoaiQuyetDinhRange");
            wb.Names.Add("DanhSachLoaiQuyetDinhRange", lookup.Cells["D2:E1000"]);

            // --- PHONG_BAN ---
            if (wb.Names.Any(n => n.Name.Equals("DanhSachPhongBan", StringComparison.OrdinalIgnoreCase))) wb.Names.Remove("DanhSachPhongBan");
            if (depts.Any()) wb.Names.Add("DanhSachPhongBan", lookup.Cells[2, 6, 1 + depts.Count, 6]);

            if (wb.Names.Any(n => n.Name.Equals("DanhSachPhongBanRange", StringComparison.OrdinalIgnoreCase))) wb.Names.Remove("DanhSachPhongBanRange");
            wb.Names.Add("DanhSachPhongBanRange", lookup.Cells["F2:G1000"]);

            // --- CHUC_DANH ---
            if (wb.Names.Any(n => n.Name.Equals("DanhSachChucDanh", StringComparison.OrdinalIgnoreCase))) wb.Names.Remove("DanhSachChucDanh");
            if (positions.Any()) wb.Names.Add("DanhSachChucDanh", lookup.Cells[2, 8, 1 + positions.Count, 8]);

            if (wb.Names.Any(n => n.Name.Equals("DanhSachChucDanhRange", StringComparison.OrdinalIgnoreCase))) wb.Names.Remove("DanhSachChucDanhRange");
            wb.Names.Add("DanhSachChucDanhRange", lookup.Cells["H2:I1000"]);

            // Thêm dữ liệu cho Thang, Ngạch, Bậc lương
            var salaryScales = AppDataContext.SalaryScales.OrderBy(s => s.Name).ToList();
            for (int i = 0; i < salaryScales.Count; i++)
            {
                lookup.Cells[i + 2, 10].Value = salaryScales[i].Name;
                lookup.Cells[i + 2, 11].Value = salaryScales[i].Id;
            }

            var salaryGrades = AppDataContext.SalaryGrades.OrderBy(g => g.Name).ToList();
            for (int i = 0; i < salaryGrades.Count; i++)
            {
                lookup.Cells[i + 2, 12].Value = salaryGrades[i].Name;
                lookup.Cells[i + 2, 13].Value = salaryGrades[i].Id;
                lookup.Cells[i + 2, 14].Value = salaryGrades[i].PaSalaryScaleId;
            }

            var salaryLevels = AppDataContext.SalaryLevels.OrderBy(l => l.Name).ToList();
            for (int i = 0; i < salaryLevels.Count; i++)
            {
                lookup.Cells[i + 2, 15].Value = salaryLevels[i].Name;
                lookup.Cells[i + 2, 16].Value = salaryLevels[i].Id;
                lookup.Cells[i + 2, 17].Value = salaryLevels[i].PaSalaryGradeId;
            }

            // Named ranges cho Thang, Ngạch, Bậc lương
            if (wb.Names.Any(n => n.Name.Equals("ThangLuong", StringComparison.OrdinalIgnoreCase))) wb.Names.Remove("ThangLuong");
            if (salaryScales.Any()) wb.Names.Add("ThangLuong", lookup.Cells[2, 10, 1 + salaryScales.Count, 10]);
            if (wb.Names.Any(n => n.Name.Equals("ThangLuongRange", StringComparison.OrdinalIgnoreCase))) wb.Names.Remove("ThangLuongRange");
            wb.Names.Add("ThangLuongRange", lookup.Cells["J2:K1000"]);

            // Tạo Named Range động cho Ngạch lương và Bậc lương
            // Lưu trữ vào các cột riêng trong sheet Data_Lookup
            // Cột 20: Name ngạch, Cột 21: Id ngạch, Cột 22: ScaleId
            // Cột 23: Name bậc, Cột 24: Id bậc, Cột 25: GradeId
            int gradeCol = 20;
            int levelCol = 23;
            var scaleGrades = salaryGrades.GroupBy(g => g.PaSalaryScaleId).ToList();
            int gradeRow = 0;
            foreach (var group in scaleGrades)
            {
                var scale = salaryScales.FirstOrDefault(s => s.Id == group.Key);
                if (scale == null) continue;

                var grades = group.OrderBy(g => g.Name).ToList();
                if (!grades.Any()) continue;

                int startRow = 2 + gradeRow;
                foreach (var grade in grades)
                {
                    lookup.Cells[2 + gradeRow, gradeCol].Value = grade.Name;
                    lookup.Cells[2 + gradeRow, gradeCol + 1].Value = grade.Id;
                    lookup.Cells[2 + gradeRow, gradeCol + 2].Value = grade.PaSalaryScaleId;
                    gradeRow++;
                }
                int endRow = startRow + grades.Count - 1;

                // Tạo named range cho INDIRECT dropdown (Ngach_{TenThang})
                var rangeName = "Ngach_" + scale.Name.Replace(" ", "_");
                if (wb.Names.Any(n => n.Name.Equals(rangeName, StringComparison.OrdinalIgnoreCase))) wb.Names.Remove(rangeName);
                // Named range chỉ chứa cột Name (cột đầu tiên) để dropdown hiển thị đúng
                wb.Names.Add(rangeName, lookup.Cells[startRow, gradeCol, endRow, gradeCol]);
            }

            // Lưu trữ Bậc lương
            var gradeLevels = salaryLevels.GroupBy(l => l.PaSalaryGradeId).ToList();
            int levelRow = 0;
            foreach (var group in gradeLevels)
            {
                var grade = salaryGrades.FirstOrDefault(g => g.Id == group.Key);
                if (grade == null) continue;

                var levels = group.OrderBy(l => l.Name).ToList();
                if (!levels.Any()) continue;

                int startRow = 2 + levelRow;
                foreach (var level in levels)
                {
                    lookup.Cells[2 + levelRow, levelCol].Value = level.Name;
                    lookup.Cells[2 + levelRow, levelCol + 1].Value = level.Id;
                    lookup.Cells[2 + levelRow, levelCol + 2].Value = level.PaSalaryGradeId;
                    levelRow++;
                }
                int endRow = startRow + levels.Count - 1;

                // Tạo named range cho INDIRECT dropdown (Bac_{TenNgach})
                var rangeName = "Bac_" + grade.Name.Replace(" ", "_");
                if (wb.Names.Any(n => n.Name.Equals(rangeName, StringComparison.OrdinalIgnoreCase))) wb.Names.Remove(rangeName);
                wb.Names.Add(rangeName, lookup.Cells[startRow, levelCol, endRow, levelCol]);
            }

            if (wb.Names.Any(n => n.Name.Equals("NgachLuongAll", StringComparison.OrdinalIgnoreCase))) wb.Names.Remove("NgachLuongAll");
            if (gradeRow > 0) wb.Names.Add("NgachLuongAll", lookup.Cells[2, gradeCol, 1 + gradeRow, gradeCol + 2]);

            if (wb.Names.Any(n => n.Name.Equals("BacLuongAll", StringComparison.OrdinalIgnoreCase))) wb.Names.Remove("BacLuongAll");
            if (levelRow > 0) wb.Names.Add("BacLuongAll", lookup.Cells[2, levelCol, 1 + levelRow, levelCol + 2]);

            var dv = ws.Cells[6, 2, 100, 2].DataValidation.AddListDataValidation();
            dv.Formula.ExcelFormula = "=DanhSachCode";
            dv.AllowBlank = true;
            dv.ErrorTitle = "Giá trị không hợp lệ";
            dv.Error = "Vui lòng chọn giá trị có trong danh sách.";

            // Dropdown Loại Quyết Định ở cột D
            var dvDecision = ws.Cells[6, 4, 100, 4].DataValidation.AddListDataValidation();
            dvDecision.Formula.ExcelFormula = "=DanhSachLoaiQuyetDinh";
            dvDecision.AllowBlank = true;
            dvDecision.ErrorTitle = "Giá trị không hợp lệ";
            dvDecision.Error = "Vui lòng chọn giá trị có trong danh sách.";

            // Dropdown Phòng Ban ở cột I
            var dvDept = ws.Cells[6, 9, 100, 9].DataValidation.AddListDataValidation();
            dvDept.Formula.ExcelFormula = "=DanhSachPhongBan";
            dvDept.AllowBlank = true;

            // Dropdown Chức Danh ở cột J
            var dvPos = ws.Cells[6, 10, 100, 10].DataValidation.AddListDataValidation();
            dvPos.Formula.ExcelFormula = "=DanhSachChucDanh";
            dvPos.AllowBlank = true;

            // Dropdown cho Thang, Ngạch, Bậc lương
            var dvScale = ws.Cells[6, 12, 100, 12].DataValidation.AddListDataValidation();
            dvScale.Formula.ExcelFormula = "=ThangLuong";
            dvScale.AllowBlank = true;

            // Dropdown Ngạch lương - phải set formula cho từng row vì cần relative reference
            for (int r = 6; r <= 100; r++)
            {
                var dvGrade = ws.Cells[r, 13, r, 13].DataValidation.AddListDataValidation();
                dvGrade.Formula.ExcelFormula = $"=IF(L{r}<>\"\",INDIRECT(\"Ngach_\" & SUBSTITUTE(L{r},\" \",\"_\")),\"\")";
                dvGrade.AllowBlank = true;

                var dvLevel = ws.Cells[r, 14, r, 14].DataValidation.AddListDataValidation();
                dvLevel.Formula.ExcelFormula = $"=IF(M{r}<>\"\",INDIRECT(\"Bac_\" & SUBSTITUTE(M{r},\" \",\"_\")),\"\")";
                dvLevel.AllowBlank = true;
            }

            for (int r = 6; r <= 100; r++)
            {
                // Cột A (1): ID nhân viên ẩn từ dropdown Mã NV (cột B)
                ws.Cells[r, 1].Formula = $"=IF(B{r}=\"\", \"\", VLOOKUP(B{r}, DanhSachCodeId, 3, FALSE))";
                // Cột C (3): Tên nhân viên từ dropdown Mã NV (cột B)
                ws.Cells[r, 3].Formula = $"=IF(B{r}=\"\", \"\", VLOOKUP(B{r}, DanhSachNhanVien, 2, FALSE))";

                // Các cột ẩn chứa ID bắt đầu từ cột Z (26) để không ảnh hưởng cột hiển thị
                // Cột 26 (Z): ID Loại QĐ ẩn từ dropdown (cột D)
                ws.Cells[r, 26].Formula = $"=IF(D{r}=\"\", \"\", VLOOKUP(D{r}, DanhSachLoaiQuyetDinhRange, 2, FALSE))";
                // Cột 27 (AA): ID Phòng ban ẩn từ dropdown (cột I)
                ws.Cells[r, 27].Formula = $"=IF(I{r}=\"\", \"\", VLOOKUP(I{r}, DanhSachPhongBanRange, 2, FALSE))";
                // Cột 28 (AB): ID Chức danh ẩn từ dropdown (cột J)
                ws.Cells[r, 28].Formula = $"=IF(J{r}=\"\", \"\", VLOOKUP(J{r}, DanhSachChucDanhRange, 2, FALSE))";
                // Cột 29 (AC): ID Thang lương ẩn từ dropdown (cột L)
                ws.Cells[r, 29].Formula = $"=IF(L{r}=\"\", \"\", VLOOKUP(L{r}, ThangLuongRange, 2, FALSE))";
                // Cột 30 (AD): ID Ngạch lương ẩn từ dropdown (cột M)
                ws.Cells[r, 30].Formula = $"=IF(M{r}=\"\", \"\", VLOOKUP(M{r}, NgachLuongAll, 2, FALSE))";
                // Cột 31 (AE): ID Bậc lương ẩn từ dropdown (cột N)
                ws.Cells[r, 31].Formula = $"=IF(N{r}=\"\", \"\", VLOOKUP(N{r}, BacLuongAll, 2, FALSE))";
                // Cột 32 (AF): ID Người ký ẩn từ dropdown (cột P)
                ws.Cells[r, 32].Formula = $"=IF(P{r}=\"\", \"\", VLOOKUP(P{r}, DanhSachCodeId, 3, FALSE))";

                // Cột Q (17): Tên người ký từ dropdown Mã người ký (cột P)
                ws.Cells[r, 17].Formula = $"=IF(P{r}=\"\", \"\", VLOOKUP(P{r}, DanhSachNhanVien, 2, FALSE))";
            }
            // Ẩn các cột ID
            ws.Column(1).Hidden = true;   // A: ID Nhân viên
            ws.Column(26).Hidden = true;  // Z: ID Loại QĐ
            ws.Column(27).Hidden = true;  // AA: ID Phòng ban
            ws.Column(28).Hidden = true;  // AB: ID Chức danh
            ws.Column(29).Hidden = true;  // AC: ID Thang lương
            ws.Column(30).Hidden = true;  // AD: ID Ngạch lương
            ws.Column(31).Hidden = true;  // AE: ID Bậc lương
            ws.Column(32).Hidden = true;  // AF: ID Người ký

            // Dropdown Mã người ký ở cột P
            var dvSigner = ws.Cells[6, 16, 100, 16].DataValidation.AddListDataValidation();
            dvSigner.Formula.ExcelFormula = "=DanhSachCode";
            dvSigner.AllowBlank = true;
            dvSigner.ErrorTitle = "Giá trị không hợp lệ";
            dvSigner.Error = "Vui lòng chọn giá trị có trong danh sách.";

            var toRemove = ws.DataValidations
                .Where(v => v.Address.Start.Column >= 5 && v.Address.End.Column <= 8)
                .ToList();
            foreach (var v in toRemove)
            {
                ws.DataValidations.Remove(v);
            }

            return File(package.GetAsByteArray(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                $"HuWorking_{DateTime.Now:yyyyMMdd}.xlsx");
        }
    }

    [HttpPost("import")]
    public IActionResult Import(IFormFile file)
    {
        if (file == null || file.Length == 0) return BadRequest("File không hợp lệ");

        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        using (var package = new ExcelPackage(file.OpenReadStream()))
        {
            var ws = package.Workbook.Worksheets[0];

            // Load data từ DB để lookup
            var map = AppDataContext.Employees
                .Join(AppDataContext.EmployeeCvs, e => e.EmployeeId, c => c.Id,
                    (e, c) => new { e.Code, EmpId = e.Id, CvId = e.EmployeeId, c.FullName })
                .ToDictionary(x => x.Code, x => x);

            var employees = new List<object>();
            var errors = new List<string>();

            DateTime? ParseDate(object v)
            {
                if (v == null) return null;
                if (v is DateTime dt) return dt;
                if (v is double d)
                {
                    try { return DateTime.FromOADate(d); } catch { }
                }
                return DateTime.TryParse(v.ToString(), out var r) ? r : (DateTime?)null;
            }

            if (ws.Dimension != null)
            {
                for (int row = 6; row <= ws.Dimension.End.Row; row++)
                {
                    // Cột B: Mã nhân viên
                    var code = ws.Cells[row, 2].Value?.ToString()?.Trim();
                    if (string.IsNullOrEmpty(code)) continue;

                    // Cột A: ID nhân viên (ẩn)
                    var employeeIdCell = ws.Cells[row, 1].Value;
                    long? employeeId = null;
                    if (employeeIdCell != null && long.TryParse(employeeIdCell.ToString(), out var empId))
                    {
                        employeeId = empId;
                    }

                    // Cột D: Loại quyết định
                    var decisionName = ws.Cells[row, 4].Value?.ToString()?.Trim();
                    long? decisionId = null;
                    var decisionIdCell = ws.Cells[row, 26].Value;  // Cột Z (ẩn)
                    if (decisionIdCell != null && long.TryParse(decisionIdCell.ToString(), out var dId))
                    {
                        decisionId = dId;
                    }

                    // Cột E: Số quyết định
                    var decisionNo = ws.Cells[row, 5].Value?.ToString()?.Trim();
                    // Cột F: Ngày hiệu lực
                    var effectiveDate = ParseDate(ws.Cells[row, 6].Value);
                    // Cột G: Ngày hết hiệu lực
                    var expireDate = ParseDate(ws.Cells[row, 7].Value);
                    // Cột H: Căn cứ QĐ số
                    var decisionBaseNo = ws.Cells[row, 8].Value?.ToString()?.Trim();
                    // Cột O: Ngày ký
                    var signedDate = ParseDate(ws.Cells[row, 15].Value);
                    // Cột R: Ghi chú
                    var note = ws.Cells[row, 18].Value?.ToString()?.Trim();

                    // Cột I: Phòng ban
                    var departmentName = ws.Cells[row, 9].Value?.ToString()?.Trim();
                    long? departmentId = null;
                    var departmentIdCell = ws.Cells[row, 27].Value;  // Cột AA (ẩn)
                    if (departmentIdCell != null && long.TryParse(departmentIdCell.ToString(), out var depId)) departmentId = depId;

                    // Cột J: Chức danh
                    var positionName = ws.Cells[row, 10].Value?.ToString()?.Trim();
                    long? positionId = null;
                    var positionIdCell = ws.Cells[row, 28].Value;  // Cột AB (ẩn)
                    if (positionIdCell != null && long.TryParse(positionIdCell.ToString(), out var posId)) positionId = posId;

                    // Cột L: Thang lương - ID từ cột ẩn 29 (AC)
                    var scaleIdCell = ws.Cells[row, 29].Value;
                    long? scaleId = null;
                    if (scaleIdCell != null && long.TryParse(scaleIdCell.ToString(), out var scId))
                    {
                        scaleId = scId;
                    }

                    // Cột M: Ngạch lương - ID từ cột ẩn 30 (AD)
                    var gradeIdCell = ws.Cells[row, 30].Value;
                    long? gradeId = null;
                    if (gradeIdCell != null && long.TryParse(gradeIdCell.ToString(), out var grId))
                    {
                        gradeId = grId;
                    }

                    // Cột N: Bậc lương - ID từ cột ẩn 31 (AE)
                    var levelIdCell = ws.Cells[row, 31].Value;
                    long? levelId = null;
                    if (levelIdCell != null && long.TryParse(levelIdCell.ToString(), out var lvId))
                    {
                        levelId = lvId;
                    }

                    // Cột P: Mã người ký
                    var signerCode = ws.Cells[row, 16].Value?.ToString()?.Trim();
                    long? signerId = null;
                    var signerIdCell = ws.Cells[row, 32].Value;  // Cột AF (ẩn)
                    if (signerIdCell != null && long.TryParse(signerIdCell.ToString(), out var sigId))
                    {
                        signerId = sigId;
                    }

                    if (string.IsNullOrEmpty(decisionNo)) errors.Add($"Dòng {row}: Thiếu số quyết định");
                    if (effectiveDate == null) errors.Add($"Dòng {row}: Thiếu ngày hiệu lực hoặc sai định dạng");

                    if (map.TryGetValue(code, out var info))
                    {
                        employees.Add(new
                        {
                            Row = row,
                            Code = code,
                            EmployeeId = employeeId ?? info.EmpId,  // Ưu tiên ID từ Excel, fallback từ DB
                            EmployeeCvId = info.CvId,
                            FullName = info.FullName,
                            DecisionTypeName = decisionName,
                            DecisionTypeId = decisionId,
                            DecisionNo = decisionNo,
                            EffectiveDate = effectiveDate,
                            ExpireDate = expireDate,
                            DecisionBaseNo = decisionBaseNo,
                            SignedDate = signedDate,
                            SignerCode = signerCode,
                            SignerId = signerId,
                            DepartmentName = departmentName,
                            DepartmentId = departmentId,
                            PositionName = positionName,
                            PositionId = positionId,
                            SalaryScaleId = scaleId,
                            SalaryGradeId = gradeId,
                            SalaryLevelId = levelId
                        });
                    }
                    else
                    {
                        errors.Add($"Dòng {row}: Mã '{code}' không tồn tại");
                    }
                }
            }

            return Ok(new
            {
                success = errors.Count == 0, totalRows = employees.Count + errors.Count, validCount = employees.Count,
                employees, errors
            });
        }
    }


}

//--Thang lương
// Các class Entity đã được chuyển ra file riêng trong thư mục Entities
//với 3 class PaSalaryScale, PaSalaryGrade, PaSalaryLevel
// Mỗi //--Thang lương có nhiều //--Ngạch lương , mỗi //--Ngạch lương có nhiều //--Bậc lương
