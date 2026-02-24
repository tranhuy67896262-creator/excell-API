using ImportExportExcellApi.Data;
using Microsoft.AspNetCore.Http;
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
        //viết 1 func dẽ đọc file excell base và sẽ thực hiện add công thức và gen ra file tên là tên file + realase in ptoduct và đọc file đó fill data
        // sau đó mới export ra file chứa data mẫu chứa các dropdow , các dropdown khi import sẽ lấy được là ID , add công thức các thứ các cột , có cột hơi đặc biệt với tôi là file

        // 1. Khởi tạo dữ liệu mẫu
        AppDataContext.Initialize();

        // 2. Query cơ bản
        var allEmployees = AppDataContext.Employees.ToList();
        Console.WriteLine($"Tổng số nhân viên: {allEmployees.Count}");

        // 3. Query với điều kiện (WHERE)
        var hanoiEmployees = AppDataContext.Employees
            .Where(e => e.Address == "Hà Nội")
            .ToList();

        // 4. Query với sắp xếp (ORDER BY)
        var sortedByAge = AppDataContext.Employees
            .OrderBy(e => e.Age)
            .ToList();

        // 5. Query với JOIN (Employee + Allowance)
        var employeeWithAllowance = from emp in AppDataContext.Employees
                                    join allow in AppDataContext.Allowances
                                    on emp.AllowanceId equals allow.Id
                                    select new
                                    {
                                        emp.Code,
                                        emp.FullName,
                                        allow.Name as AllowanceName
                                    };
    }
}
