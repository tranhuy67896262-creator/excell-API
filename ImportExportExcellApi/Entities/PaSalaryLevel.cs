namespace ImportExportExcellApi.Entities
{
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
}