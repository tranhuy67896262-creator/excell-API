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
        private static List<EmployeeCv> _employeeCvs = new List<EmployeeCv>();
        private static List<SysOtherList> _sysOtherLists = new List<SysOtherList>();
        private static List<Department> _departments = new List<Department>();
        private static List<Position> _positions = new List<Position>();
        private static List<PaSalaryScale> _salaryScales = new List<PaSalaryScale>();
        private static List<PaSalaryGrade> _salaryGrades = new List<PaSalaryGrade>();
        private static List<PaSalaryLevel> _salaryLevels = new List<PaSalaryLevel>();

        // Properties để query với LINQ
        public static IQueryable<Allowance> Allowances => _allowances.AsQueryable();
        public static IQueryable<Employee> Employees => _employees.AsQueryable();
        public static IQueryable<EmployeeCv> EmployeeCvs => _employeeCvs.AsQueryable();
        public static IQueryable<SysOtherList> SysOtherLists => _sysOtherLists.AsQueryable();
        public static IQueryable<Department> Departments => _departments.AsQueryable();
        public static IQueryable<Position> Positions => _positions.AsQueryable();
        public static IQueryable<PaSalaryScale> SalaryScales => _salaryScales.AsQueryable();
        public static IQueryable<PaSalaryGrade> SalaryGrades => _salaryGrades.AsQueryable();
        public static IQueryable<PaSalaryLevel> SalaryLevels => _salaryLevels.AsQueryable();

        /// <summary>
        /// Khởi tạo dữ liệu mẫu (chạy 1 lần khi khởi động)
        /// </summary>
        public static void Initialize()
        {
            // Nếu đã có dữ liệu rồi thì không init lại
            if (_employees.Any() || _employeeCvs.Any())
                return;

            // 1. Init dữ liệu cho EmployeeCv (Hồ sơ chi tiết)
            var cv1 = new EmployeeCv { Id = 11, FullName = "Nguyễn Văn A", Age = 30, Address = "Hà Nội", Email = "a@example.com", Phone = "0901111111", DateBirth = new DateTime(1996, 5, 20) };
            var cv2 = new EmployeeCv { Id = 22, FullName = "Trần Thị B", Age = 28, Address = "TP.HCM", Email = "b@example.com", Phone = "0902222222", DateBirth = new DateTime(1998, 8, 15) };
            var cv3 = new EmployeeCv { Id = 32, FullName = "Lê Văn C", Age = 35, Address = "Đà Nẵng", Email = "c@example.com", Phone = "0903333333", DateBirth = new DateTime(1991, 2, 10) };
            var cv4 = new EmployeeCv { Id = 44, FullName = "Phạm Thị D", Age = 25, Address = "Hải Phòng", Email = "d@example.com", Phone = "0904444444", DateBirth = new DateTime(2001, 11, 5) };

            _employeeCvs.AddRange(new[] { cv1, cv2, cv3, cv4 });

            // 2. Init dữ liệu cho Employee (Bảng liên kết, chứa Code và ID tham chiếu)
            // Lưu ý: EmployeeId trỏ đến Id của EmployeeCv
            _employees.AddRange(new[]
            {
                new Employee { Id = 1, Code = "EMP001", EmployeeId = 11 }, // Nguyễn Văn A
                new Employee { Id = 2, Code = "EMP002", EmployeeId = 22 }, // Trần Thị B
                new Employee { Id = 3, Code = "EMP003", EmployeeId = 33 }, // Lê Văn C
                new Employee { Id = 4, Code = "EMP004", EmployeeId = 44 }  // Phạm Thị D
            });

            // 3. Init dữ liệu cho SysOtherList (Dùng cho Dropdown khác)
            _sysOtherLists.AddRange(new[]
            {
                new SysOtherList { Id = 31, Name = "Loại bằng cấp/Chứng chỉ 1" ,TypeCode = "CERTIFICATE_TYPE"},
                new SysOtherList { Id = 311, Name = "Loại bằng cấp/Chứng chỉ 2" ,TypeCode = "CERTIFICATE_TYPE"},
                new SysOtherList { Id = 312, Name = "Loại bằng cấp/Chứng chỉ 3" ,TypeCode = "CERTIFICATE_TYPE"},

                new SysOtherList { Id = 134, Name = "Đơn vị đào tạo 1",TypeCode= "GRADUATE_SCHOOL" },
                new SysOtherList { Id = 234, Name = "Đơn vị đào tạo 2",TypeCode= "GRADUATE_SCHOOL" },

                new SysOtherList { Id = 42, Name = "Trình độ chuyên môn 1" ,TypeCode = "LEVEL_ID"},
                new SysOtherList { Id = 43, Name = "Trình độ chuyên môn 2" ,TypeCode = "LEVEL_ID"},
                new SysOtherList { Id = 44, Name = "Trình độ chuyên môn 3" ,TypeCode = "LEVEL_ID"},
                //LEARNING_LEVEL  trình độ học vấn
                new SysOtherList { Id = 52, Name = "Trình độ học vấn 1" ,TypeCode = "LEVEL_TRAIN"},
                new SysOtherList { Id = 53, Name = "Trình độ học vấn 2" ,TypeCode = "LEVEL_TRAIN"},
                new SysOtherList { Id = 54, Name = "Trình độ học vấn 3" ,TypeCode = "LEVEL_TRAIN"},
                //
                new SysOtherList { Id = 62, Name = "Hình thức đào tạo 1" ,TypeCode = "TRAINING_METHOD"},
                new SysOtherList { Id = 63, Name = "Hình thức đào tạo 2" ,TypeCode = "TRAINING_METHOD"},
                new SysOtherList { Id = 64, Name = "Hình thức đào tạo 3" ,TypeCode = "TRAINING_METHOD"},

                // TYPE_DECISION - Loại quyết định
                new SysOtherList { Id = 201, Name = "Quyết định bổ nhiệm" ,TypeCode = "TYPE_DECISION"},
                new SysOtherList { Id = 202, Name = "Quyết định điều chuyển" ,TypeCode = "TYPE_DECISION"},
                new SysOtherList { Id = 203, Name = "Quyết định khen thưởng" ,TypeCode = "TYPE_DECISION"},
                new SysOtherList { Id = 204, Name = "Quyết định kỷ luật" ,TypeCode = "TYPE_DECISION"},

            });

            // 4. Init Departments
            _departments.AddRange(new[]
            {
                new Department { Id = 1, Code = "PB01", Name = "Phòng Nhân sự" },
                new Department { Id = 2, Code = "PB02", Name = "Phòng Kỹ thuật" },
                new Department { Id = 3, Code = "PB03", Name = "Phòng Kinh doanh" },
                new Department { Id = 4, Code = "PB04", Name = "Phòng Kế toán" }
            });

            // 5. Init Positions
            _positions.AddRange(new[]
            {
                new Position { Id = 1, Code = "CD01", Name = "Trưởng phòng" },
                new Position { Id = 2, Code = "CD02", Name = "Phó phòng" },
                new Position { Id = 3, Code = "CD03", Name = "Chuyên viên" },
                new Position { Id = 4, Code = "CD04", Name = "Nhân viên" }
            });

            // 6. Init Salary Structures
            _salaryScales.AddRange(new[] {
                new PaSalaryScale { Id = 1, Code = "TS01", Name = "Thang Lương Văn Phòng" },
                new PaSalaryScale { Id = 2, Code = "TS02", Name = "Thang Lương Kỹ Thuật" }
            });

            _salaryGrades.AddRange(new[] {
                new PaSalaryGrade { Id = 11, PaSalaryScaleId = 1, Code = "NG01", Name = "Ngạch Chuyên Viên" },
                new PaSalaryGrade { Id = 12, PaSalaryScaleId = 1, Code = "NG02", Name = "Ngạch Quản Lý" },
                new PaSalaryGrade { Id = 21, PaSalaryScaleId = 2, Code = "NG03", Name = "Ngạch Kỹ Sư" }
            });

            _salaryLevels.AddRange(new[] {
                new PaSalaryLevel { Id = 111, PaSalaryGradeId = 11, Code = "B01", Name = "Bậc 1" },
                new PaSalaryLevel { Id = 112, PaSalaryGradeId = 11, Code = "B02", Name = "Bậc 2" },
                new PaSalaryLevel { Id = 121, PaSalaryGradeId = 12, Code = "B01", Name = "Bậc 1" },
                new PaSalaryLevel { Id = 211, PaSalaryGradeId = 21, Code = "B01", Name = "Bậc 1" }
            });

            // 7. Init dữ liệu cho Allowance (Giả định class này vẫn tồn tại để tránh lỗi)
            // Nếu class Allowance đã bị xóa, bạn có thể bỏ đoạn này và xóa các reference liên quan
            _allowances.AddRange(new[]
            {
                new Allowance { Id = 1, Name = "Phụ cấp ăn trưa", Amount = 730000 },
                new Allowance { Id = 2, Name = "Phụ cấp xăng xe", Amount = 1000000 },
                new Allowance { Id = 3, Name = "Phụ cấp điện thoại", Amount = 300000 }
            });
        }

        /// <summary>
        /// Xóa toàn bộ dữ liệu (dùng cho testing)
        /// </summary>
        public static void Clear()
        {
            _allowances.Clear();
            _employees.Clear();
            _employeeCvs.Clear();
            _sysOtherLists.Clear();
        }

        #region Employee Methods

        public static Employee AddEmployee(Employee employee)
        {
            employee.Id = _employees.Any() ? _employees.Max(e => e.Id) + 1 : 1;
            _employees.Add(employee);
            return employee;
        }

        public static bool UpdateEmployee(Employee employee)
        {
            var existing = _employees.FirstOrDefault(e => e.Id == employee.Id);
            if (existing == null) return false;

            // Cập nhật thông tin cơ bản của bảng Employee
            existing.Code = employee.Code;

            // Nếu cần cập nhật cả thông tin chi tiết (CV) thì phải tìm theo EmployeeId
            if (employee.EmployeeId > 0)
            {
                var cvToUpdate = _employeeCvs.FirstOrDefault(c => c.Id == employee.EmployeeId);
                if (cvToUpdate != null)
                {
                    // Lưu ý: Class Employee hiện tại không chứa FullName, Age... 
                    // nên việc update các trường này phải truyền qua một object EmployeeCv riêng hoặc sửa logic gọi hàm.
                    // Ở đây tôi chỉ cập nhật liên kết ID.
                    existing.EmployeeId = employee.EmployeeId;
                }
            }

            return true;
        }

        public static bool DeleteEmployee(long id)
        {
            var employee = _employees.FirstOrDefault(e => e.Id == id);
            if (employee == null) return false;

            _employees.Remove(employee);
            // Lưu ý: Có thể xóa cả CV liên quan nếu muốn, tùy nghiệp vụ
            // var cv = _employeeCvs.FirstOrDefault(c => c.Id == employee.EmployeeId);
            // if(cv != null) _employeeCvs.Remove(cv);

            return true;
        }

        public static Employee GetEmployeeById(long id)
        {
            return _employees.FirstOrDefault(e => e.Id == id);
        }

        /// <summary>
        /// Lấy thông tin đầy đủ của Employee kèm theo CV
        /// </summary>
        public static dynamic GetEmployeeWithCv(long id)
        {
            var emp = _employees.FirstOrDefault(e => e.Id == id);
            if (emp == null) return null;

            var cv = _employeeCvs.FirstOrDefault(c => c.Id == emp.EmployeeId);

            return new
            {
                emp.Id,
                emp.Code,
                cv.FullName,
                cv.Age,
                cv.Email,
                cv.Phone,
                cv.Address,
                cv.DateBirth
            };
        }

        #endregion

        #region EmployeeCv Methods

        public static EmployeeCv AddEmployeeCv(EmployeeCv cv)
        {
            cv.Id = _employeeCvs.Any() ? _employeeCvs.Max(c => c.Id) + 1 : 1;
            _employeeCvs.Add(cv);
            return cv;
        }

        public static bool UpdateEmployeeCv(EmployeeCv cv)
        {
            var existing = _employeeCvs.FirstOrDefault(c => c.Id == cv.Id);
            if (existing == null) return false;

            existing.FullName = cv.FullName;
            existing.Age = cv.Age;
            existing.Address = cv.Address;
            existing.Email = cv.Email;
            existing.Phone = cv.Phone;
            existing.DateBirth = cv.DateBirth;
            return true;
        }

        public static EmployeeCv GetEmployeeCvById(long id)
        {
            return _employeeCvs.FirstOrDefault(c => c.Id == id);
        }

        #endregion

        #region SysOtherList Methods

        public static SysOtherList AddSysOtherList(SysOtherList item)
        {
            item.Id = _sysOtherLists.Any() ? _sysOtherLists.Max(i => i.Id) + 1 : 1;
            _sysOtherLists.Add(item);
            return item;
        }

        public static List<SysOtherList> GetAllSysOtherLists()
        {
            return _sysOtherLists.ToList();
        }

        #endregion

        #region Allowance Methods (Giữ lại để tương thích)

        public static Allowance GetAllowanceById(long id)
        {
            return _allowances.FirstOrDefault(a => a.Id == id);
        }

        public static List<Allowance> GetAllAllowances()
        {
            return _allowances.ToList();
        }

        #endregion
    }
}