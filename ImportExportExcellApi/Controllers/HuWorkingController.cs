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

    [HttpGet("generate-template")]
    public IActionResult GenerateTemplate()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        var basePath = Path.Combine(_environment.ContentRootPath, _templateFolderName, _baseTemplateName);
        var releasePath = Path.Combine(_environment.ContentRootPath, _templateFolderName, _releaseTemplateName);

        using (var package = new ExcelPackage(new FileInfo(basePath)))
        {
            var ws = package.Workbook.Worksheets[0];
            var lookup = package.Workbook.Worksheets.FirstOrDefault(w => w.Name == "Data_Lookup") ??
                         package.Workbook.Worksheets.Add("Data_Lookup");

            lookup.Hidden = eWorkSheetHidden.VeryHidden;
            lookup.Cells.Clear();

            var wb = package.Workbook;

            // 1. Thiết lập Named Ranges "Tĩnh" trỏ vào các vùng dữ liệu trong Data_Lookup
            
            // --- Nhân viên (Cột A, B, C) ---
            if (wb.Names.Any(n => n.Name.Equals("DanhSachCode", StringComparison.OrdinalIgnoreCase))) wb.Names.Remove("DanhSachCode");
            wb.Names.Add("DanhSachCode", lookup.Cells["A2:A2000"]);
            
            if (wb.Names.Any(n => n.Name.Equals("DanhSachNhanVien", StringComparison.OrdinalIgnoreCase))) wb.Names.Remove("DanhSachNhanVien");
            wb.Names.Add("DanhSachNhanVien", lookup.Cells["A2:B2000"]);

            if (wb.Names.Any(n => n.Name.Equals("DanhSachCodeId", StringComparison.OrdinalIgnoreCase))) wb.Names.Remove("DanhSachCodeId");
            wb.Names.Add("DanhSachCodeId", lookup.Cells["A2:C2000"]);

            // --- Loại quyết định (Cột D, E) ---
            if (wb.Names.Any(n => n.Name.Equals("DanhSachLoaiQuyetDinh", StringComparison.OrdinalIgnoreCase))) wb.Names.Remove("DanhSachLoaiQuyetDinh");
            wb.Names.Add("DanhSachLoaiQuyetDinh", lookup.Cells["D2:D1000"]);
            if (wb.Names.Any(n => n.Name.Equals("DanhSachLoaiQuyetDinhRange", StringComparison.OrdinalIgnoreCase))) wb.Names.Remove("DanhSachLoaiQuyetDinhRange");
            wb.Names.Add("DanhSachLoaiQuyetDinhRange", lookup.Cells["D2:E1000"]);

            // --- Phòng ban (Cột F, G) ---
            if (wb.Names.Any(n => n.Name.Equals("DanhSachPhongBan", StringComparison.OrdinalIgnoreCase))) wb.Names.Remove("DanhSachPhongBan");
            wb.Names.Add("DanhSachPhongBan", lookup.Cells["F2:F1000"]);
            if (wb.Names.Any(n => n.Name.Equals("DanhSachPhongBanRange", StringComparison.OrdinalIgnoreCase))) wb.Names.Remove("DanhSachPhongBanRange");
            wb.Names.Add("DanhSachPhongBanRange", lookup.Cells["F2:G1000"]);

            // --- Chức danh (Cột H, I) ---
            if (wb.Names.Any(n => n.Name.Equals("DanhSachChucDanh", StringComparison.OrdinalIgnoreCase))) wb.Names.Remove("DanhSachChucDanh");
            wb.Names.Add("DanhSachChucDanh", lookup.Cells["H2:H1000"]);
            if (wb.Names.Any(n => n.Name.Equals("DanhSachChucDanhRange", StringComparison.OrdinalIgnoreCase))) wb.Names.Remove("DanhSachChucDanhRange");
            wb.Names.Add("DanhSachChucDanhRange", lookup.Cells["H2:I1000"]);

            // --- Thang lương (Cột J, K, L) ---
            if (wb.Names.Any(n => n.Name.Equals("ThangLuongList", StringComparison.OrdinalIgnoreCase))) wb.Names.Remove("ThangLuongList");
            wb.Names.Add("ThangLuongList", lookup.Cells["J2:J1000"]);
            if (wb.Names.Any(n => n.Name.Equals("ThangLuongData", StringComparison.OrdinalIgnoreCase))) wb.Names.Remove("ThangLuongData");
            wb.Names.Add("ThangLuongData", lookup.Cells["J2:L1000"]); 

            // --- Ngạch lương (Cột M, N, O, P) ---
            if (wb.Names.Any(n => n.Name.Equals("NgachLuongData", StringComparison.OrdinalIgnoreCase))) wb.Names.Remove("NgachLuongData");
            wb.Names.Add("NgachLuongData", lookup.Cells["M2:P1000"]);

            // --- Bậc lương (Cột R, S, T) ---
            if (wb.Names.Any(n => n.Name.Equals("BacLuongData", StringComparison.OrdinalIgnoreCase))) wb.Names.Remove("BacLuongData");
            wb.Names.Add("BacLuongData", lookup.Cells["R2:S1000"]); // Name -> Id

            // 2. Thiết lập Data Validation & Formulas
            var dvEmp = ws.Cells[6, 2, 1000, 2].DataValidation.AddListDataValidation();
            dvEmp.Formula.ExcelFormula = "=DanhSachCode";

            var dvDecision = ws.Cells[6, 4, 1000, 4].DataValidation.AddListDataValidation();
            dvDecision.Formula.ExcelFormula = "=DanhSachLoaiQuyetDinh";

            var dvDept = ws.Cells[6, 9, 1000, 9].DataValidation.AddListDataValidation();
            dvDept.Formula.ExcelFormula = "=DanhSachPhongBan";

            var dvPos = ws.Cells[6, 10, 1000, 10].DataValidation.AddListDataValidation();
            dvPos.Formula.ExcelFormula = "=DanhSachChucDanh";

            var dvScale = ws.Cells[6, 12, 1000, 12].DataValidation.AddListDataValidation();
            dvScale.Formula.ExcelFormula = "=ThangLuongList";

            // Dropdown Mã người ký (Cột P)
            var dvSigner = ws.Cells[6, 16, 1000, 16].DataValidation.AddListDataValidation();
            dvSigner.Formula.ExcelFormula = "=DanhSachCode";

            for (int r = 6; r <= 1000; r++)
            {
                var dvGrade = ws.Cells[r, 13, r, 13].DataValidation.AddListDataValidation();
                dvGrade.Formula.ExcelFormula = $"=IF(L{r}<>\"\",INDIRECT(VLOOKUP(L{r},ThangLuongData,3,FALSE)),\"\")";

                var dvLevel = ws.Cells[r, 14, r, 14].DataValidation.AddListDataValidation();
                dvLevel.Formula.ExcelFormula = $"=IF(M{r}<>\"\",INDIRECT(VLOOKUP(M{r},NgachLuongData,4,FALSE)),\"\")";

                ws.Cells[r, 1].Formula = $"=IF(B{r}=\"\", \"\", VLOOKUP(B{r}, DanhSachCodeId, 3, FALSE))";
                ws.Cells[r, 3].Formula = $"=IF(B{r}=\"\", \"\", VLOOKUP(B{r}, DanhSachNhanVien, 2, FALSE))";
                ws.Cells[r, 26].Formula = $"=IF(D{r}=\"\", \"\", VLOOKUP(D{r}, DanhSachLoaiQuyetDinhRange, 2, FALSE))";
                ws.Cells[r, 27].Formula = $"=IF(I{r}=\"\", \"\", VLOOKUP(I{r}, DanhSachPhongBanRange, 2, FALSE))";
                ws.Cells[r, 28].Formula = $"=IF(J{r}=\"\", \"\", VLOOKUP(J{r}, DanhSachChucDanhRange, 2, FALSE))";

                // Tên người ký (Cột Q) và ID người ký (Cột AF - 32)
                ws.Cells[r, 17].Formula = $"=IF(P{r}=\"\", \"\", VLOOKUP(P{r}, DanhSachNhanVien, 2, FALSE))";
                ws.Cells[r, 32].Formula = $"=IF(P{r}=\"\", \"\", VLOOKUP(P{r}, DanhSachCodeId, 3, FALSE))";

                // Tra cứu ID Thang/Ngạch/Bậc lương (Cột 29, 30, 31)
                ws.Cells[r, 29].Formula = $"=IF(L{r}=\"\", \"\", VLOOKUP(L{r}, ThangLuongData, 2, FALSE))";
                ws.Cells[r, 30].Formula = $"=IF(M{r}=\"\", \"\", VLOOKUP(M{r}, NgachLuongData, 2, FALSE))";
                ws.Cells[r, 31].Formula = $"=IF(N{r}=\"\", \"\", VLOOKUP(N{r}, BacLuongData, 2, FALSE))";
            }

            ws.Column(1).Hidden = true;
            ws.Column(26).Hidden = true;
            ws.Column(27).Hidden = true;
            ws.Column(28).Hidden = true;
            ws.Column(29).Hidden = true;
            ws.Column(30).Hidden = true;
            ws.Column(31).Hidden = true;
            ws.Column(32).Hidden = true;

            package.SaveAs(new FileInfo(releasePath));
        }
        return Ok(new { message = "Base template generated successfully." });
    }

    [HttpGet("export")]
    public IActionResult Export()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        var releasePath = Path.Combine(_environment.ContentRootPath, _templateFolderName, _releaseTemplateName);
        
        if (!System.IO.File.Exists(releasePath)) return BadRequest("Vui lòng chạy generate-template trước.");

        using (var package = new ExcelPackage(new FileInfo(releasePath)))
        {
            var ws = package.Workbook.Worksheets[0];
            var lookup = package.Workbook.Worksheets["Data_Lookup"];
            var wb = package.Workbook;

            // Xóa dữ liệu cũ trong lookup và sheet chính để tránh bị rác
            lookup.Cells.Clear(); 
            // Xóa vùng dữ liệu nhân viên trong lookup (theo yêu cầu)
            lookup.Cells["A2:C1000"].Clear(); 

            // Xóa trắng các cột nhập liệu trong sheet chính (từ dòng 6) để không bị "thừa" dữ liệu cũ
            // Cột B (Mã NV), D->P (Các thông tin), R (Ghi chú)
            ws.Cells["B6:P1000"].Value = null;
            ws.Cells["R6:R1000"].Value = null;

            // 1. Điền dữ liệu Nhân viên
            // Cấu trúc Data_Lookup: Cột A=Code, Cột B=FullName, Cột C=Id
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

            // Định nghĩa lại Named Range cho Nhân viên để không bị thừa dòng trống
            // DanhSachCode: chỉ cột Code (cột A) - dùng cho dropdown
            if (wb.Names.Any(n => n.Name.Equals("DanhSachCode", StringComparison.OrdinalIgnoreCase))) wb.Names.Remove("DanhSachCode");
            if (emps.Any()) wb.Names.Add("DanhSachCode", lookup.Cells[2, 1, 1 + emps.Count, 1]);

            // DanhSachNhanVien: 2 cột (Code, FullName) - VLOOKUP lấy tên từ cột 2
            if (wb.Names.Any(n => n.Name.Equals("DanhSachNhanVien", StringComparison.OrdinalIgnoreCase))) wb.Names.Remove("DanhSachNhanVien");
            if (emps.Any()) wb.Names.Add("DanhSachNhanVien", lookup.Cells[2, 1, 1 + emps.Count, 2]);

            // DanhSachCodeId: 3 cột (Code, FullName, Id) - VLOOKUP lấy ID từ cột 3
            if (wb.Names.Any(n => n.Name.Equals("DanhSachCodeId", StringComparison.OrdinalIgnoreCase))) wb.Names.Remove("DanhSachCodeId");
            if (emps.Any()) wb.Names.Add("DanhSachCodeId", lookup.Cells[2, 1, 1 + emps.Count, 3]);

            // Cập nhật lại formulas trong sheet chính để đảm bảo VLOOKUP đúng
            for (int r = 6; r <= 1000; r++)
            {
                ws.Cells[r, 1].Formula = $"=IF(B{r}=\"\", \"\", VLOOKUP(B{r}, DanhSachCodeId, 3, FALSE))";
                ws.Cells[r, 3].Formula = $"=IF(B{r}=\"\", \"\", VLOOKUP(B{r}, DanhSachNhanVien, 2, FALSE))";
            }

            // 2. Điền dữ liệu Danh mục động
            var decisions = AppDataContext.SysOtherLists.Where(s => s.TypeCode == "TYPE_DECISION").OrderBy(s => s.Name).ToList();
            for (int i = 0; i < decisions.Count; i++) { lookup.Cells[i + 2, 4].Value = decisions[i].Name; lookup.Cells[i + 2, 5].Value = decisions[i].Id; }
            
            if (wb.Names.Any(n => n.Name.Equals("DanhSachLoaiQuyetDinh", StringComparison.OrdinalIgnoreCase))) wb.Names.Remove("DanhSachLoaiQuyetDinh");
            if (decisions.Any()) wb.Names.Add("DanhSachLoaiQuyetDinh", lookup.Cells[2, 4, 1 + decisions.Count, 4]);
            
            if (wb.Names.Any(n => n.Name.Equals("DanhSachLoaiQuyetDinhRange", StringComparison.OrdinalIgnoreCase))) wb.Names.Remove("DanhSachLoaiQuyetDinhRange");
            if (decisions.Any()) wb.Names.Add("DanhSachLoaiQuyetDinhRange", lookup.Cells[2, 4, 1 + decisions.Count, 5]);

            var depts = AppDataContext.Departments.OrderBy(d => d.Name).ToList();
            for (int i = 0; i < depts.Count; i++) { lookup.Cells[i + 2, 6].Value = depts[i].Name; lookup.Cells[i + 2, 7].Value = depts[i].Id; }

            if (wb.Names.Any(n => n.Name.Equals("DanhSachPhongBan", StringComparison.OrdinalIgnoreCase))) wb.Names.Remove("DanhSachPhongBan");
            if (depts.Any()) wb.Names.Add("DanhSachPhongBan", lookup.Cells[2, 6, 1 + depts.Count, 6]);

            if (wb.Names.Any(n => n.Name.Equals("DanhSachPhongBanRange", StringComparison.OrdinalIgnoreCase))) wb.Names.Remove("DanhSachPhongBanRange");
            if (depts.Any()) wb.Names.Add("DanhSachPhongBanRange", lookup.Cells[2, 6, 1 + depts.Count, 7]);

            var positions = AppDataContext.Positions.OrderBy(p => p.Name).ToList();
            for (int i = 0; i < positions.Count; i++) { lookup.Cells[i + 2, 8].Value = positions[i].Name; lookup.Cells[i + 2, 9].Value = positions[i].Id; }

            if (wb.Names.Any(n => n.Name.Equals("DanhSachChucDanh", StringComparison.OrdinalIgnoreCase))) wb.Names.Remove("DanhSachChucDanh");
            if (positions.Any()) wb.Names.Add("DanhSachChucDanh", lookup.Cells[2, 8, 1 + positions.Count, 8]);

            if (wb.Names.Any(n => n.Name.Equals("DanhSachChucDanhRange", StringComparison.OrdinalIgnoreCase))) wb.Names.Remove("DanhSachChucDanhRange");
            if (positions.Any()) wb.Names.Add("DanhSachChucDanhRange", lookup.Cells[2, 8, 1 + positions.Count, 9]);

            // 3. Điền dữ liệu Danh mục Lương (Sắp xếp theo ParentId để dữ liệu liên tục)
            var salaryScales = AppDataContext.SalaryScales.OrderBy(s => s.Name).ToList();
            for (int i = 0; i < salaryScales.Count; i++)
            {
                lookup.Cells[i + 2, 10].Value = salaryScales[i].Name;
                lookup.Cells[i + 2, 11].Value = salaryScales[i].Id;
                lookup.Cells[i + 2, 12].Value = "SCALE_" + salaryScales[i].Code.Replace(" ", "_");
            }

            if (wb.Names.Any(n => n.Name.Equals("ThangLuongList", StringComparison.OrdinalIgnoreCase))) wb.Names.Remove("ThangLuongList");
            if (salaryScales.Any()) wb.Names.Add("ThangLuongList", lookup.Cells[2, 10, 1 + salaryScales.Count, 10]);

            if (wb.Names.Any(n => n.Name.Equals("ThangLuongData", StringComparison.OrdinalIgnoreCase))) wb.Names.Remove("ThangLuongData");
            if (salaryScales.Any()) wb.Names.Add("ThangLuongData", lookup.Cells[2, 10, 1 + salaryScales.Count, 12]);

            // QUAN TRỌNG: Sắp xếp theo ScaleId để các ngạch cùng thang nằm cạnh nhau
            var salaryGrades = AppDataContext.SalaryGrades.OrderBy(g => g.PaSalaryScaleId).ThenBy(g => g.Name).ToList();
            for (int i = 0; i < salaryGrades.Count; i++)
            {
                lookup.Cells[i + 2, 13].Value = salaryGrades[i].Name;
                lookup.Cells[i + 2, 14].Value = salaryGrades[i].Id;
                lookup.Cells[i + 2, 15].Value = salaryGrades[i].PaSalaryScaleId;
                lookup.Cells[i + 2, 16].Value = "GRADE_" + salaryGrades[i].Code.Replace(" ", "_");
            }

            if (wb.Names.Any(n => n.Name.Equals("NgachLuongData", StringComparison.OrdinalIgnoreCase))) wb.Names.Remove("NgachLuongData");
            if (salaryGrades.Any()) wb.Names.Add("NgachLuongData", lookup.Cells[2, 13, 1 + salaryGrades.Count, 16]);

            // QUAN TRỌNG: Sắp xếp theo GradeId để các bậc cùng ngạch nằm cạnh nhau
            var salaryLevels = AppDataContext.SalaryLevels.OrderBy(l => l.PaSalaryGradeId).ThenBy(l => l.Name).ToList();
            for (int i = 0; i < salaryLevels.Count; i++)
            {
                lookup.Cells[i + 2, 18].Value = salaryLevels[i].Name;
                lookup.Cells[i + 2, 19].Value = salaryLevels[i].Id;
                lookup.Cells[i + 2, 20].Value = salaryLevels[i].PaSalaryGradeId;
            }

            if (wb.Names.Any(n => n.Name.Equals("BacLuongData", StringComparison.OrdinalIgnoreCase))) wb.Names.Remove("BacLuongData");
            if (salaryLevels.Any()) wb.Names.Add("BacLuongData", lookup.Cells[2, 18, 1 + salaryLevels.Count, 19]);

            // Tạo Named Range động cho Ngạch lương và Bậc lương dựa trên dữ liệu vừa nạp
            foreach (var scale in salaryScales)
            {
                var gradesOfScale = salaryGrades.Where(g => g.PaSalaryScaleId == scale.Id).ToList();
                if (gradesOfScale.Any())
                {
                    var rangeName = "SCALE_" + scale.Code.Replace(" ", "_");
                    if (wb.Names.Any(n => n.Name.Equals(rangeName, StringComparison.OrdinalIgnoreCase))) wb.Names.Remove(rangeName);
                    
                    var firstGrade = gradesOfScale.First();
                    var startRow = 2 + salaryGrades.IndexOf(firstGrade);
                    wb.Names.Add(rangeName, lookup.Cells[startRow, 13, startRow + gradesOfScale.Count - 1, 13]);
                }
            }

            foreach (var grade in salaryGrades)
            {
                var levelsOfGrade = salaryLevels.Where(l => l.PaSalaryGradeId == grade.Id).ToList();
                if (levelsOfGrade.Any())
                {
                    var rangeName = "GRADE_" + grade.Code.Replace(" ", "_");
                    if (wb.Names.Any(n => n.Name.Equals(rangeName, StringComparison.OrdinalIgnoreCase))) wb.Names.Remove(rangeName);

                    var firstLevel = levelsOfGrade.First();
                    var startRow = 2 + salaryLevels.IndexOf(firstLevel);
                    wb.Names.Add(rangeName, lookup.Cells[startRow, 18, startRow + levelsOfGrade.Count - 1, 18]);
                }
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
                            EmployeeId = employeeId ?? info.EmpId,
                            EmployeeCvId = info.CvId,
                            FullName = info.FullName,
                            DecisionTypeName = decisionName,
                            DecisionTypeId = decisionId,
                            DecisionNo = decisionNo,
                            EffectiveDate = effectiveDate,
                            ExpireDate = expireDate,
                            DecisionBaseNo = decisionBaseNo,
                            SignedDate = signedDate,
                            Note = note,
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
