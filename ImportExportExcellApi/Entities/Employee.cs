using System.ComponentModel.DataAnnotations;

namespace ImportExportExcellApi.Entities
{
    public class Employee
    {
        public long Id { get; set; }
        public string Code { get; set; }
        public long EmployeeId { get; set; }// là ID của bảng EmployeeCv
    }
}
