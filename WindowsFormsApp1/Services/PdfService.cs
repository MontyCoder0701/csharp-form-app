using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using PdfLib;

namespace WindowsFormsApp1.Services
{
    internal class PdfEmployeeData
    {
        public string name;
        public string uidnum7;
        public string baseYear;
        public decimal preCalculatedSalary;
        public decimal deductibleTax;

        public PdfEmployeeData(string name, string uidnum7, string baseYear, decimal preCalculatedSalary, decimal deductibleTax)
        {
            this.name = name;
            this.uidnum7 = uidnum7;
            this.baseYear = baseYear;
            this.preCalculatedSalary = preCalculatedSalary;
            this.deductibleTax = deductibleTax;
        }
    }

    internal class PdfService
    {
        private static PdfEmployeeData Extract2024PdfEmployeeDataFromFile(string filePath)
        {
            List<List<string>> firstTableData = PdfManager.ImportPdfToTable(filePath, 1);

            firstTableData = firstTableData
               .Select(row => row.Select(cell => Regex.Replace(cell, @"\s+", "")).ToList())
               .ToList();

            string name = firstTableData
                  .FirstOrDefault(row => row.Any(cell => cell.Contains("⑥") && row.Any(c => c.Contains("⑦"))))?
                  .SkipWhile(cell => !cell.Contains("⑥"))
                  .Skip(2)
                  .FirstOrDefault() ?? "";

            string uid = firstTableData
                .FirstOrDefault(row => row.Any(cell => cell.Contains("⑥") && row.Any(c => c.Contains("⑦"))))?
                .Last()
                .Substring(0, 7)
                .Replace("-", "") ?? "";

            if (name.Length == 0 || uid.Length == 0)
            {
                throw new Exception("잘못된 파일입니다.");
            }

            if (name.Contains("*") || uid.Contains("*"))
            {
                throw new Exception("이름이나 주민번호가 가려져 직원 정보를 알 수 없습니다.");
            }

            string baseYear = firstTableData
                .FirstOrDefault(row => row.Any(cell => cell.Contains("근무기간")))?
                .SkipWhile(cell => !cell.Contains("근무기간"))
                .Skip(1)
                .FirstOrDefault()
                .Substring(0, 4) ?? "";

            if (string.IsNullOrWhiteSpace(baseYear) || !int.TryParse(baseYear, out int year))
            {
                throw new Exception("잘못된 파일입니다.");
            }

            if (year != DateTime.Now.Year - 1)
            {
                throw new Exception("작년 기준년도의 원천징수영수증이 아닙니다.");
            }

            decimal totalSum = firstTableData
               .FirstOrDefault(row => row.Any(cell => (cell.Contains("16") && cell.Contains("계")) || cell == "계"))?
               .SkipWhile(cell => !(cell.Contains("16") && cell.Contains("계")) && cell != "계")
               .Skip(1)
               .Select(cell => decimal.TryParse(cell.Trim(), out decimal value) ? value : 0)
               .FirstOrDefault() ?? 0;

            decimal untaxedTotalSum = firstTableData
                .FirstOrDefault(row => row.Any(cell => cell.Contains("비과세소득")))?
                .SkipWhile(cell => !cell.Contains("비과세소득"))
                .Skip(1)
                .Select(cell => decimal.TryParse(cell.Trim(), out decimal value) ? value : 0)
                .FirstOrDefault() ?? 0;

            decimal previousTaxPaid = firstTableData
               .FirstOrDefault(row => row.Any(cell => cell.Contains("주(현)근무지")))?
               .SkipWhile(cell => !cell.Contains("주(현)근무지"))
               .Skip(1)
               .Take(3)
               .Select(cell => decimal.TryParse(cell.Trim(), out decimal value) ? value : 0)
               .Sum() ?? 0;

            decimal deductibleTax = firstTableData
               .FirstOrDefault(row => row.Any(cell => cell.Contains("징수세액")))?
               .SkipWhile(cell => !cell.Contains("징수세액"))
               .Skip(1)
               .Take(3)
               .Select(cell => decimal.TryParse(cell.Trim(), out decimal value) ? value : 0)
               .Sum() ?? 0;

            List<List<string>> secondTableData = PdfManager.ImportPdfToTable(filePath, 2, 3, 25);

            secondTableData = secondTableData
               .Select(row => row.Select(cell => Regex.Replace(cell, @"\s+", "")).ToList())
               .ToList();

            int nationalPensionRowIndex = secondTableData
                    .FindIndex(row => row.Any(cell => cell.Contains("국민연금보험료")));

            decimal nationalPension = nationalPensionRowIndex < 0 ? 0 : secondTableData[nationalPensionRowIndex - 1]
                .SkipWhile(cell => cell != "대상금액")
                .Skip(1)
                .Select(cell => decimal.TryParse(cell.Trim(), out decimal value) ? value : 0)
                .FirstOrDefault();

            int publicOfficialPensionIndex = secondTableData
                .FindIndex(row => row.Any(cell => cell.Contains("공무원연금")));

            decimal publicOfficialPension = publicOfficialPensionIndex < 0 ? 0 : secondTableData[publicOfficialPensionIndex - 1]
                .SkipWhile(cell => cell != "대상금액")
                .Skip(1)
                .Select(cell => decimal.TryParse(cell.Trim(), out decimal value) ? value : 0)
                .FirstOrDefault();

            int soldierPensionIndex = secondTableData
                .FindIndex(row => row.Any(cell => cell.Contains("군인연금")));

            decimal soldierPension = secondTableData
              .Where((row, index) => index == soldierPensionIndex || index == soldierPensionIndex - 1)
              .SelectMany(row => row.SkipWhile(cell => cell != "대상금액")
              .Skip(1))
              .Select(cell => decimal.TryParse(cell.Trim(), out decimal value) ? value : (decimal?)null)
              .FirstOrDefault() ?? 0;

            int privateSchoolPensionRowIndex = secondTableData
                .FindIndex(row => row.Any(cell => cell.Contains("사립학교")));

            decimal privateSchoolPension = secondTableData
               .Where((row, index) => index == privateSchoolPensionRowIndex || index == privateSchoolPensionRowIndex - 1)
               .SelectMany(row => row.SkipWhile(cell => cell != "대상금액")
               .Skip(1))
               .Select(cell => decimal.TryParse(cell.Trim(), out decimal value) ? value : (decimal?)null)
               .FirstOrDefault() ?? 0;

            int postalPensionRowIndex = secondTableData
                .FindIndex(row => row.Any(cell => cell.Contains("별정우체국")));

            decimal postalPension = postalPensionRowIndex < 0 ? 0 : secondTableData[postalPensionRowIndex - 1]
                .SkipWhile(cell => cell != "대상금액")
                .Skip(1)
                .Select(cell => decimal.TryParse(cell.Trim(), out decimal value) ? value : 0)
                .FirstOrDefault();

            int healthInsuranceIndex = secondTableData
                .FindIndex(row => row.Any(cell => cell.Contains("건강보험료")));

            decimal healthInsurance = healthInsuranceIndex < 0 ? 0 : secondTableData[healthInsuranceIndex]
                .SkipWhile(cell => cell != "대상금액")
                .Skip(1)
                .Select(cell => decimal.TryParse(cell.Trim(), out decimal value) ? value : 0)
                .FirstOrDefault();

            int employmentInsuranceIndex = secondTableData
                .FindIndex(row => row.Any(cell => cell.Contains("고용보험료")));

            decimal employmentInsurance = employmentInsuranceIndex < 0 ? 0 : secondTableData[employmentInsuranceIndex - 1]
                .SkipWhile(cell => cell != "대상금액")
                .Skip(1)
                .Select(cell => decimal.TryParse(cell.Trim(), out decimal value) ? value : 0)
                .FirstOrDefault();

            decimal preCalculatedSalary = totalSum + untaxedTotalSum - previousTaxPaid - (nationalPension + publicOfficialPension + soldierPension + privateSchoolPension + postalPension + healthInsurance + employmentInsurance);

            return new PdfEmployeeData(name, uid, baseYear, preCalculatedSalary, deductibleTax);
        }

        public static List<PdfEmployeeData> ExtractPdfEmployeeDataFromFiles(List<string> filePaths)
        {
            List<PdfEmployeeData> employeeDataList = new List<PdfEmployeeData>();

            foreach (var filePath in filePaths)
            {
                try
                {
                    // 매년 해당 함수만 바꾸면 됌
                    PdfEmployeeData pdfEmployeeData = Extract2024PdfEmployeeDataFromFile(filePath);
                    employeeDataList.Add(pdfEmployeeData);
                }
                catch
                {
                    continue;
                }

            }

            return employeeDataList;
        }
    }
}
