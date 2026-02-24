namespace ImportExportExcellApi.Dto
{
    public class TrainingRecordDto
    {
        public string Code { get; set; }           // Mã nhân viên (String)
        public long EmployeeCvId { get; set; }     // ID hồ sơ (Long) - Tương ứng FULL_NAME
        public string FullName { get; set; }       // Tên hiển thị (String) - Để fill vào ô dropdown

        public long CertTypeId { get; set; }       // ID loại bằng (Long) - Tương ứng CERTIFICATE_TYPE
        public string CertTypeName { get; set; }   // Tên loại bằng (String)

        public bool IsPrime { get; set; }          // Là bằng chính (Boolean/Yes-No)

        public string Name { get; set; }           // Tên bằng cấp (String)

        public long LevelId { get; set; }          // ID trình độ (Long) - Tương ứng LEVEL_ID
        public string LevelName { get; set; }      // Tên trình độ (String)

        public string LevelTrain { get; set; }     // Trình độ học vấn (String)
        public string Method { get; set; }         // Hình thức đào tạo (String)

        public int Year { get; set; }              // Năm (Int)

        public string Content { get; set; }        // Nội dung (String)
        public decimal Mark { get; set; }          // Điểm số (Decimal)

        public DateTime TrainFrom { get; set; }
        public DateTime TrainTo { get; set; }
        public DateTime EffectFrom { get; set; }
        public DateTime EffectTo { get; set; }

        public string Classification { get; set; } // Xếp loại (String)
        public string Remark { get; set; }         // Ghi chú (String)
    }
}
