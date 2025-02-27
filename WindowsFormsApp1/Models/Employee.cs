using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using Newtonsoft.Json;

namespace WindowsFormsApp1.Models
{
    internal class Employee
    {
        public int EmplSeq { get; }
        public string ProjCode { get; }
        public string Cid { get; }
        public string EmplName { get; set; }
        public string Uidnum7 { get; }
        public string EmplNum { get; set; }

        private int _salaryBcode;

        public int SalaryBcode
        {
            get => _salaryBcode;
            set
            {
                _salaryBcode = value;
                SalaryBname = bCodeToName.TryGetValue(value, out var name) ? name : "";
            }
        }

        public string SalaryBname { get; private set; }
        public string SalaryAcctnum { get; set; }
        public int SalaryAmt { get; set; }
        public int SalaryBaseYear { get; set; }
        public int? DeductibleTax { get; set; }
        public int? DeductibleTaxBaseYear { get; set; }
        public string EmploymentDate { get; set; }
        public int IsSafe { get; set; }
        public string Registdate { get; }
        public string Registdateformat { get; }

        public string GetDisplayUidnum7 => Uidnum7.Substring(0, 6) + "-" + Uidnum7.Substring(6, 1) + "******";

        public string GetDisplayIsSafe => IsSafe == 1 ? "O" : "X";

        public static readonly Dictionary<int, string> bCodeToName = new Dictionary<int, string>()
        {
            { 101, "우리" },
            { 102, "신한" },
            { 103, "국민" },
            { 104, "하나" },
            { 105, "농협" }
        };

        public Employee(
            int emplSeq,
            string cid,
            string emplName,
            string uidnum7,
            string emplNum,
            int salaryBcode,
            string salaryAcctnum,
            int salaryAmt,
            int salaryBaseYear,
            string employmentDate,
            int isSafe,
            string registdate,
            string registdateformat
        )
        {
            EmplSeq = emplSeq;
            ProjCode = "GOTCHOO";
            Cid = cid;
            EmplName = emplName;
            Uidnum7 = uidnum7;
            EmplNum = emplNum;
            SalaryBcode = salaryBcode;
            SalaryBname = bCodeToName.TryGetValue(salaryBcode, out var name) ? name : "";
            SalaryAcctnum = salaryAcctnum;
            SalaryAmt = salaryAmt;
            SalaryBaseYear = salaryBaseYear;
            EmploymentDate = employmentDate;
            IsSafe = isSafe;
            Registdate = registdate;
            Registdateformat = registdateformat;
        }
    }

    internal class EmployeeExcelDto
    {
        [Required(ErrorMessage = "직원 이름을 입력해야 합니다.")]
        public string EmplName { get; set; }

        [Required(ErrorMessage = "주민등록번호 (앞 7자리)를 입력해야 합니다.")]
        [MinLength(7, ErrorMessage = "주민등록번호는 최소 7자리를 입력해야 합니다.")]
        public string GetDisplayUidnum7 { get; set; }

        [JsonIgnore]
        public string Uidnum7 => GetDisplayUidnum7.Trim().Replace("-", "").Substring(0, 7);

        [Required(ErrorMessage = "직원 번호를 입력해야 합니다.")]
        public string EmplNum { get; set; }

        [JsonIgnore]
        public int SalaryBcode => Employee.bCodeToName.FirstOrDefault(x => x.Value == SalaryBname.Trim()).Key;

        [Required(ErrorMessage = "은행 이름을 입력해야 합니다.")]
        public string SalaryBname { get; set; }

        [Required(ErrorMessage = "계좌 번호를 입력해야 합니다.")]
        public string SalaryAcctnum { get; set; }

        [Required(ErrorMessage = "급여 금액을 입력해야 합니다.")]
        [Range(1, int.MaxValue, ErrorMessage = "급여 금액은 1원 이상이어야 합니다.")]
        public int SalaryAmt { get; set; }

        [Required(ErrorMessage = "기준 연도를 입력해야 합니다.")]
        [Range(1900, 2100, ErrorMessage = "기준 연도는 1900에서 2100 사이여야 합니다.")]
        public int SalaryBaseYear { get; set; }

        [Required(ErrorMessage = "고용 날짜를 입력해야 합니다.")]
        [DataType(DataType.Date, ErrorMessage = "고용 날짜는 올바른 날짜 형식이어야 합니다.")]
        public string EmploymentDate { get; set; }

        [Required(ErrorMessage = "안전 여부를 입력해야 합니다.")]
        [RegularExpression(@"[OX]", ErrorMessage = "안전 여부는 'O' 또는 'X'여야 합니다.")]
        public string GetDisplayIsSafe { get; set; }

        [JsonIgnore]
        public int IsSafe => GetDisplayIsSafe.Trim() == "O" ? 1 : 0;

        public string Validate()
        {
            var results = new List<ValidationResult>();
            var context = new ValidationContext(this);
            bool isValid = Validator.TryValidateObject(this, context, results, true);

            return isValid ? null : results.First().ErrorMessage;
        }
    }
}
