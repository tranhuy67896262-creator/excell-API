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

            var wb = package.Workbook;
            if (wb.Names.Any(n => n.Name.Equals("DanhSachCode", StringComparison.OrdinalIgnoreCase)))
                wb.Names.Remove("DanhSachCode");
            wb.Names.AddFormula("DanhSachCode", "OFFSET('Data_Lookup'!$A$2,0,0,COUNTA('Data_Lookup'!$A:$A)-1,1)");

            if (wb.Names.Any(n => n.Name.Equals("DanhSachNhanVien", StringComparison.OrdinalIgnoreCase)))
                wb.Names.Remove("DanhSachNhanVien");
            wb.Names.Add("DanhSachNhanVien", lookup.Cells["A2:B1000"]);

            if (wb.Names.Any(n => n.Name.Equals("DanhSachCodeId", StringComparison.OrdinalIgnoreCase)))
                wb.Names.Remove("DanhSachCodeId");
            wb.Names.Add("DanhSachCodeId", lookup.Cells["A2:C1000"]);

            // TYPE_DECISION ranges
            if (wb.Names.Any(n => n.Name.Equals("DanhSachLoaiQuyetDinh", StringComparison.OrdinalIgnoreCase)))
                wb.Names.Remove("DanhSachLoaiQuyetDinh");
            wb.Names.AddFormula("DanhSachLoaiQuyetDinh",
                "OFFSET('Data_Lookup'!$D$2,0,0,COUNTA('Data_Lookup'!$D:$D)-1,1)");
            if (wb.Names.Any(n => n.Name.Equals("DanhSachLoaiQuyetDinhRange", StringComparison.OrdinalIgnoreCase)))
                wb.Names.Remove("DanhSachLoaiQuyetDinhRange");
            wb.Names.Add("DanhSachLoaiQuyetDinhRange", lookup.Cells["D2:E1000"]);

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

            for (int r = 6; r <= 100; r++)
            {
                ws.Cells[r, 3].Formula = $"=IF(B{r}=\"\", \"\", VLOOKUP(B{r}, DanhSachNhanVien, 2, FALSE))";
                ws.Cells[r, 1].Formula = $"=IF(B{r}=\"\", \"\", VLOOKUP(B{r}, DanhSachCodeId, 3, FALSE))";
                ws.Cells[r, 26].Formula = $"=IF(D{r}=\"\", \"\", VLOOKUP(D{r}, DanhSachLoaiQuyetDinhRange, 2, FALSE))";
            }

            ws.Column(26).Hidden = true;

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
            var map = AppDataContext.Employees
                .Join(AppDataContext.EmployeeCvs, e => e.EmployeeId, c => c.Id,
                    (e, c) => new { e.Code, EmpId = e.Id, CvId = e.EmployeeId, c.FullName })
                .ToDictionary(x => x.Code, x => x);

            var employees = new List<object>();
            var errors = new List<string>();

            if (ws.Dimension != null)
            {
                for (int row = 6; row <= ws.Dimension.End.Row; row++)
                {
                    var code = ws.Cells[row, 2].Value?.ToString()?.Trim();
                    if (string.IsNullOrEmpty(code)) continue;

                    var decisionName = ws.Cells[row, 4].Value?.ToString()?.Trim();
                    long? decisionId = null;
                    var decisionIdCell = ws.Cells[row, 26].Value;
                    if (decisionIdCell != null && long.TryParse(decisionIdCell.ToString(), out var dId))
                    {
                        decisionId = dId;
                    }
                    else if (!string.IsNullOrEmpty(decisionName))
                    {
                        errors.Add($"Dòng {row}: Loại quyết định '{decisionName}' không hợp lệ");
                    }

                    if (map.TryGetValue(code, out var info))
                    {
                        employees.Add(new
                        {
                            Row = row, Code = code, EmployeeId = info.EmpId, EmployeeCvId = info.CvId,
                            FullName = info.FullName, DecisionTypeName = decisionName, DecisionTypeId = decisionId
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

    [HttpPost("test-import-file-with-data")]
    public IActionResult ImportFileWithData()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        return Ok(new { message = "Release Template đã sẵn sàng" });
    }
}

//--Thang lương
// Các class Entity nên để ở file riêng, nhưng để đây vẫn đúng cú pháp
public class PaSalaryScale
{
    public int Id { get; set; }
    public string Code { get; set; }
    public string Name { get; set; }
    public decimal? SalaryBase { get; set; }
    public decimal? SalaryAllowance { get; set; }
    public decimal? SalaryBonus { get; set; }
    public decimal? SalaryPenalty { get; set; }
    public decimal? SalaryTotal { get; set; }
    public virtual ICollection<PaSalaryGrade> SalaryGrades { get; set; }
}

//--Ngạch lương
public class PaSalaryGrade
{
    public int Id { get; set; }
    public string Code { get; set; }
    public string Name { get; set; }
    public int PaSalaryScaleId { get; set; }
    public decimal? SalaryBase { get; set; }
    public decimal? SalaryAllowance { get; set; }
    public decimal? SalaryBonus { get; set; }
    public decimal? SalaryPenalty { get; set; }
    public decimal? SalaryTotal { get; set; }

    public virtual PaSalaryScale SalaryScale { get; set; }
    public virtual ICollection<PaSalaryLevel> SalaryLevels { get; set; }
}

//-- Bậc lương
public class PaSalaryLevel
{
    public int Id { get; set; }
    public string Code { get; set; }
    public string Name { get; set; }
    public int PaSalaryGradeId { get; set; }
    public decimal? SalaryBase { get; set; }
    public decimal? SalaryAllowance { get; set; }
    public decimal? SalaryBonus { get; set; }
    public decimal? SalaryPenalty { get; set; }
    public decimal? SalaryTotal { get; set; }

    public virtual PaSalaryGrade SalaryGrade { get; set; }
}

//với 3 class PaSalaryScale, PaSalaryGrade, PaSalaryLevel
// Mỗi //--Thang lương có nhiều //--Ngạch lương , mỗi //--Ngạch lương có nhiều //--Bậc lương
