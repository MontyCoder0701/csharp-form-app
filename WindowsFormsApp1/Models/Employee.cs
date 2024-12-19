namespace WindowsFormsApp1.Models
{
    internal class Employee
    {
        public int EmplSeq { get; }
        public string ProjCode { get; }
        public string Cid { get; }
        public string EmplName { get; }
        public string Uidnum7 { get; }
        public string HealthInsuranceDate { get; }
        public string EmplNum { get; }
        public string Grade { get; }
        public string Department { get; }
        public int SalaryBcode { get; }
        public string SalaryBname { get; }
        public string SalaryAcctnum { get; }
        public int AllowanceBcode { get; }
        public string AllowanceBname { get; }
        public string AllowanceAcctnum { get; }
        public int SalaryAmt { get; }
        public int SalaryBaseYear { get; }
        public int DeductibleTax { get; }
        public int DeductibleTaxBaseYear { get; }
        public string EmploymentDate { get; }
        public int IsSafe { get; }
        public string Registdate { get; }
        public string Registdateformat { get; }
        public string Memo { get; }

        public Employee(
            int emplSeq,
            string projCode,
            string cid,
            string emplName,
            string uidnum7,
            string healthInsuranceDate,
            string emplNum,
            string grade,
            string department,
            int salaryBcode,
            string salaryBname,
            string salaryAcctnum,
            int allowanceBcode,
            string allowanceBname,
            string allowanceAcctnum,
            int salaryAmt,
            int salaryBaseYear,
            int deductibleTax,
            int deductibleTaxBaseYear,
            string employmentDate,
            int isSafe,
            string registdate,
            string registdateformat,
            string memo
        )
        {
            EmplSeq = emplSeq;
            ProjCode = projCode;
            Cid = cid;
            EmplName = emplName;
            Uidnum7 = uidnum7;
            HealthInsuranceDate = healthInsuranceDate;
            EmplNum = emplNum;
            Grade = grade;
            Department = department;
            SalaryBcode = salaryBcode;
            SalaryBname = salaryBname;
            SalaryAcctnum = salaryAcctnum;
            AllowanceBcode = allowanceBcode;
            AllowanceBname = allowanceBname;
            AllowanceAcctnum = allowanceAcctnum;
            SalaryAmt = salaryAmt;
            SalaryBaseYear = salaryBaseYear;
            DeductibleTax = deductibleTax;
            DeductibleTaxBaseYear = deductibleTaxBaseYear;
            EmploymentDate = employmentDate;
            IsSafe = isSafe;
            Registdate = registdate;
            Registdateformat = registdateformat;
            Memo = memo;
        }
    }

}
