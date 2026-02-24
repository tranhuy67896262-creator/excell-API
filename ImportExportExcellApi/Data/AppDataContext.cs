using System;
using System.Collections.Generic;
using System.Linq;
using ImportExportExcellApi.Entities;

namespace ImportExportExcellApi.Data
{
    /// <summary>
    /// DataContext giả lập để lưu trữ dữ liệu trong memory
    /// Hỗ trợ LINQ như database thật
    /// </summary>
    public class AppDataContext
    {
        // Collections lưu trữ dữ liệu
        private static List<Allowance> _allowances = new List<Allowance>();
        private static List<Employee> _employees = new List<Employee>();

        // Properties để query với LINQ
        public static IQueryable<Allowance> Allowances => _allowances.AsQueryable();
        public static IQueryable<Employee> Employees => _employees.AsQueryable();

        /// <summary>
        /// Khởi tạo dữ liệu mẫu (chạy 1 lần khi khởi động)
        /// </summary>
        public static void Initialize()
        {
            // Nếu đã có dữ liệu rồi thì không init lại
            if (_allowances.Any() || _employees.Any())
                return;

            // 1. Khởi tạo Allowances
            _allowances = new List<Allowance>
            {
                new Allowance { Id = 11, Code = "ALLOW001", Name = "Phụ cấp ăn trưa" },
                new Allowance { Id = 22, Code = "ALLOW002", Name = "Phụ cấp xăng xe" },
                new Allowance { Id = 33, Code = "ALLOW003", Name = "Phụ cấp điện thoại" },
                new Allowance { Id = 44, Code = "ALLOW004", Name = "Phụ cấp nhà ở" },
                new Allowance { Id = 55, Code = "ALLOW005", Name = "Không có phụ cấp" }
            };

            // 2. Khởi tạo Employees
            _employees = new List<Employee>
            {
                new Employee
                {
                    Id = 1,
                    Code = "EMP001",
                    FullName = "Nguyễn Văn A",
                    Age = 25,
                    Address = "Hà Nội",
                    Email = "nguyenvana@email.com",
                    Phone = "0901234567",
                    DateBirth = new DateTime(1998, 5, 15),
                    AllowanceId = 11
                },
                new Employee
                {
                    Id = 2,
                    Code = "EMP002",
                    FullName = "Trần Thị B",
                    Age = 30,
                    Address = "TP.HCM",
                    Email = "tranthib@email.com",
                    Phone = "0912345678",
                    DateBirth = new DateTime(1993, 8, 20),
                    AllowanceId = 22
                },
                new Employee
                {
                    Id = 3,
                    Code = "EMP003",
                    FullName = "Lê Văn C",
                    Age = 28,
                    Address = "Đà Nẵng",
                    Email = "levanc@email.com",
                    Phone = "0923456789",
                    DateBirth = new DateTime(1995, 12, 10),
                    AllowanceId = 33
                },
                new Employee
                {
                    Id = 4,
                    Code = "EMP004",
                    FullName = "Phạm Thị D",
                    Age = 35,
                    Address = "Hải Phòng",
                    Email = "phamthid@email.com",
                    Phone = "0934567890",
                    DateBirth = new DateTime(1988, 3, 5),
                    AllowanceId = 11
                },
                new Employee
                {
                    Id = 5,
                    Code = "EMP005",
                    FullName = "Hoàng Văn E",
                    Age = 22,
                    Address = "Cần Thơ",
                    Email = "hoangvane@email.com",
                    Phone = "0945678901",
                    DateBirth = new DateTime(2001, 7, 25),
                    AllowanceId = 55
                }
            };
        }

        /// <summary>
        /// Xóa toàn bộ dữ liệu (dùng cho testing)
        /// </summary>
        public static void Clear()
        {
            _allowances.Clear();
            _employees.Clear();
        }

        /// <summary>
        /// Thêm mới Employee
        /// </summary>
        public static Employee AddEmployee(Employee employee)
        {
            employee.Id = _employees.Any() ? _employees.Max(e => e.Id) + 1 : 1;
            _employees.Add(employee);
            return employee;
        }

        /// <summary>
        /// Cập nhật Employee
        /// </summary>
        public static bool UpdateEmployee(Employee employee)
        {
            var existing = _employees.FirstOrDefault(e => e.Id == employee.Id);
            if (existing == null) return false;

            existing.Code = employee.Code;
            existing.FullName = employee.FullName;
            existing.Age = employee.Age;
            existing.Address = employee.Address;
            existing.Email = employee.Email;
            existing.Phone = employee.Phone;
            existing.DateBirth = employee.DateBirth;
            existing.AllowanceId = employee.AllowanceId;

            return true;
        }

        /// <summary>
        /// Xóa Employee theo Id
        /// </summary>
        public static bool DeleteEmployee(long id)
        {
            var employee = _employees.FirstOrDefault(e => e.Id == id);
            if (employee == null) return false;

            _employees.Remove(employee);
            return true;
        }

        /// <summary>
        /// Tìm Employee theo Id
        /// </summary>
        public static Employee GetEmployeeById(long id)
        {
            return _employees.FirstOrDefault(e => e.Id == id);
        }

        /// <summary>
        /// Tìm Allowance theo Id
        /// </summary>
        public static Allowance GetAllowanceById(long id)
        {
            return _allowances.FirstOrDefault(a => a.Id == id);
        }
    }
}