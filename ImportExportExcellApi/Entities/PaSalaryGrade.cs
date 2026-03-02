using System.Collections.Generic;

namespace ImportExportExcellApi.Entities
{
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
}