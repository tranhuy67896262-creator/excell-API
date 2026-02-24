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
            string releaseTemplatePath = Path.Combine(templateFolder, _releaseTemplateName);

            if (!System.IO.File.Exists(baseTemplatePath))
            {
                return NotFound(new { message = $"Không tìm thấy file Base Template tại: {baseTemplatePath}" });
            }

            GenerateReleaseTemplate(baseTemplatePath, releaseTemplatePath);

            var sampleData = new[]
            {
                new {
                    EmployeeId = 1211L,
                    Code = "EMP001",
                    FullName = "Nguyễn Văn A",
                    certificateType = 98L,
                    certificateTypeName = "Bằng Đại học",
                    IsPrime = "Có",
                },
                new {
                    EmployeeId = 222L,
                    Code = "EMP002",
                    FullName = "Trần Thị B",
                    certificateType = 99L,
                    certificateTypeName = "Bằng cao đẳng",
                    IsPrime = "Không",
                }
            };

            byte[] fileBytes;
            string fileName = $"Mau_Nhap_Lieu_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";

            using (var package = new ExcelPackage(new FileInfo(releaseTemplatePath)))
            {
                var ws = package.Workbook.Worksheets[0];
                int startRow = 6;

                foreach (var item in sampleData)
                {
                    int row = startRow++;
                    ws.Cells[row, 2].Value = item.Code;
                    ws.Cells[row, 3].Value = item.FullName;
                    ws.Cells[row, 4].Value = item.EmployeeId;
                    ws.Cells[row, 5].Value = item.certificateTypeName;
                    ws.Cells[row, 6].Value = item.certificateType;
                    ws.Cells[row, 7].Value = item.IsPrime;
                }

                fileBytes = package.GetAsByteArray();
            }

            return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        }

        private void GenerateReleaseTemplate(string sourcePath, string destPath)
        {
            using (var package = new ExcelPackage(new FileInfo(sourcePath)))
            {
                var ws = package.Workbook.Worksheets[0];
                string dsSheetName = "DataSources";

                if (package.Workbook.Worksheets.Any(x => x.Name == dsSheetName))
                {
                    package.Workbook.Worksheets.Delete(dsSheetName);
                }

                var dsWs = package.Workbook.Worksheets.Add(dsSheetName);
                dsWs.Hidden = eWorkSheetHidden.Hidden;

                var employees = AppDataContext.EmployeeCvs.ToList();
                var certificates = AppDataContext.SysOtherLists.ToList();

                for (int i = 0; i < employees.Count; i++)
                {
                    dsWs.Cells[i + 1, 1].Value = employees[i].Id;
                    dsWs.Cells[i + 1, 2].Value = employees[i].FullName;
                }

                for (int i = 0; i < certificates.Count; i++)
                {
                    dsWs.Cells[i + 1, 4].Value = certificates[i].Id;
                    dsWs.Cells[i + 1, 5].Value = certificates[i].Name;
                }

                if (employees.Any())
                {
                    package.Workbook.Names.Add("List_EmpNames", dsWs.Cells[1, 2, employees.Count, 2]);
                }
                if (certificates.Any())
                {
                    package.Workbook.Names.Add("List_CertNames", dsWs.Cells[1, 5, certificates.Count, 5]);
                }

                int startRow = 6;
                int endRow = 100;

                // === XỬ LÝ FULL_NAME & ID ẨN ===
                ApplyDropdown(ws, startRow, endRow, 3, "List_EmpNames");

                for (int r = startRow; r <= endRow; r++)
                {
                    string idColumnRange = $"DataSources!$A$1:$A${employees.Count}";
                    string nameColumnRange = $"DataSources!$B$1:$B${employees.Count}";

                    ws.Cells[r, 4].Formula = $"IF(C{r}=\"\", \"\", INDEX({idColumnRange}, MATCH(C{r}, {nameColumnRange}, 0)))";
                }
                ws.Column(4).Hidden = true;

                // === XỬ LÝ CERTIFICATE_TYPE & ID ẨN ===
                ApplyDropdown(ws, startRow, endRow, 5, "List_CertNames");

                for (int r = startRow; r <= endRow; r++)
                {
                    string idColumnRange = $"DataSources!$D$1:$D${certificates.Count}";
                    string nameColumnRange = $"DataSources!$E$1:$E${certificates.Count}";

                    ws.Cells[r, 6].Formula = $"IF(E{r}=\"\", \"\", INDEX({idColumnRange}, MATCH(E{r}, {nameColumnRange}, 0)))";
                }
                ws.Column(6).Hidden = true;

                // === XỬ LÝ IS_PRIME ===
                var primeValidation = ws.Cells[startRow, 7, endRow, 7].DataValidation.AddListDataValidation();

                // SỬA: EPPlus 5.x+ dùng Formula.ExcelFormula
                primeValidation.Formula.ExcelFormula = "\"Có,Không\"";

                primeValidation.ShowErrorMessage = true;
                primeValidation.ErrorTitle = "Dữ liệu không hợp lệ";
                primeValidation.Error = "Vui lòng chọn Có hoặc Không.";

                package.SaveAs(new FileInfo(destPath));
            }
        }


        private void ApplyDropdown(ExcelWorksheet ws, int startRow, int endRow, int column, string namedRange)
        {
            if (ws.Workbook.Names.Any(n => n.Name == namedRange))
            {
                var range = ws.Cells[startRow, column, endRow, column];
                var validation = range.DataValidation.AddListDataValidation();

                // EPPlus 5.x+ dùng Formula.ExcelFormula
                validation.Formula.ExcelFormula = $"INDIRECT(\"{namedRange}\")";

                validation.ShowErrorMessage = true;
                validation.ErrorTitle = "Chọn từ danh sách";
                validation.Error = "Vui lòng chọn giá trị có sẵn.";
                validation.PromptTitle = "Hướng dẫn";
                validation.Prompt = "Nhấn vào mũi tên để chọn.";
                validation.ShowInputMessage = true;
            }
        }

    }
}