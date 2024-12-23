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
        public int SalaryBcode { get; }
        public string SalaryBname { get; set; }
        public string SalaryAcctnum { get; set; }
        public int SalaryAmt { get; set; }
        public int SalaryBaseYear { get; set; }
        public string EmploymentDate { get; set; }
        public int IsSafe { get; set; }
        public string Registdate { get; }
        public string Registdateformat { get; }

        public string GetDisplayUidnum7 => Uidnum7.Substring(0, 6) + "-" + Uidnum7.Substring(6, 1) + "******";

        public string GetDisplayIsSafe => IsSafe == 1 ? "O" : "X";


        public Employee(
            int emplSeq,
            string cid,
            string emplName,
            string uidnum7,
            string emplNum,
            int salaryBcode,
            string salaryBname,
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
            SalaryBname = salaryBname;
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
        public string EmplName { get; set; }

        public string DisplayUidnum7 { get; set; }

        [JsonIgnore]
        public string Uidnum7 => DisplayUidnum7.Replace("-", "").Substring(0, 7);

        public string EmplNum { get; set; }

        public string SalaryBname { get; set; }

        public string SalaryAcctnum { get; set; }

        public int SalaryAmt { get; set; }

        public int SalaryBaseYear { get; set; }

        public string EmploymentDate { get; set; }

        public string DisplayIsSafe { get; set; }

        [JsonIgnore]
        public int IsSafe => DisplayIsSafe == "O" ? 1 : 0;
    }
}
