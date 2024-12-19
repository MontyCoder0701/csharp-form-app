namespace WindowsFormsApp1.Models
{
    internal class Employee
    {
        public int EmplSeq { get; }
        public string ProjCode { get; }
        public string Cid { get; }
        public string EmplName { get; }
        public string Uidnum7 { get; }
        public string EmplNum { get; }
        public int SalaryBcode { get; }
        public string SalaryBname { get; }
        public string SalaryAcctnum { get; }
        public int SalaryAmt { get; }
        public int SalaryBaseYear { get; }
        public string EmploymentDate { get; }
        public int IsSafe { get; }
        public string Registdate { get; }
        public string Registdateformat { get; }

        public Employee(
            int emplSeq,
            string projCode,
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
            ProjCode = projCode;
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

}
