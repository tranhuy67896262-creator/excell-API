using System.ComponentModel.DataAnnotations;

namespace ImportExportExcellApi.Entities
{
    public class Employee
    {
        public long Id { get; set; }
        public string Code { get; set; }

        public string FullName { get; set; }

        public int Age { get; set; }

        public string Address { get; set; }


        public string Email { get; set; }


        public string Phone { get; set; }

        public DateTime DateBirth { get; set; }
        public long AllowanceId { get; set; }
    }
}
