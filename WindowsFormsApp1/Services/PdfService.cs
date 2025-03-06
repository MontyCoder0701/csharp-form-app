using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using PdfLib;

namespace WindowsFormsApp1.Services
{
    public class PdfEmployeeData
    {
        public string Name { get; }
        public string Uidnum7 { get; }
        public int BaseYear { get; }
        public int PreCalculatedSalary { get; }
        public int DeductibleTax { get; }

        public PdfEmployeeData(string name, string uidnum7, int baseYear, int preCalculatedSalary, int deductibleTax)
        {
            Name = name;
            Uidnum7 = uidnum7;
            BaseYear = baseYear;
            PreCalculatedSalary = preCalculatedSalary;
            DeductibleTax = deductibleTax;
        }
    }

    public interface IPdfEmployeeDataExtractor
    {
        PdfEmployeeData ExtractPdfEmployeeData(string filePath);
    }

    public class PdfEmployeeDataExtractor2024 : IPdfEmployeeDataExtractor
    {
        public PdfEmployeeData ExtractPdfEmployeeData(string filePath)
        {
            List<List<string>> firstTableData = PdfManager.ImportPdfToTable(filePath, 1);

            firstTableData = firstTableData
               .Select(row => row.Select(cell => Regex.Replace(cell, @"\s+", "")).ToList())
               .ToList();

            // 디버깅용
            foreach (var row in firstTableData)
            {
                Console.WriteLine(string.Join("|", row));
            }

            string name = firstTableData
                  .FirstOrDefault(row => row.Any(cell => cell.Contains("⑥") && row.Any(c => c.Contains("⑦"))))?
                  .SkipWhile(cell => !cell.Contains("⑥"))
                  .Skip(2)
                  .FirstOrDefault() ?? "";

            string uid = firstTableData
                .FirstOrDefault(row => row.Any(cell => cell.Contains("⑥") && row.Any(c => c.Contains("⑦"))))?
                .Last()
                .Substring(0, 8)
                .Replace("-", "") ?? "";

            //if (name.Length == 0 || uid.Length == 0)
            //{
            //    throw new Exception("잘못된 파일입니다.");
            //}

            //if (name.Contains("*") || uid.Contains("*"))
            //{
            //    throw new Exception("이름이나 주민번호가 가려져 직원 정보를 알 수 없습니다.");
            //}

            int baseYear = firstTableData
                .FirstOrDefault(row => row.Any(cell => cell.Contains("근무기간")))?
                .SkipWhile(cell => !cell.Contains("근무기간"))
                .Skip(1)
                .Select(cell => int.TryParse(cell.Trim().Substring(0, 4), out int value) ? value : 0)
                .FirstOrDefault() ?? 0;

            //if (baseYear == 0)
            //{
            //    throw new Exception("잘못된 파일입니다.");
            //}

            //if (baseYear != DateTime.Now.Year - 1)
            //{
            //    throw new Exception("작년 기준년도의 원천징수영수증이 아닙니다.");
            //}

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

            int deductibleTax = firstTableData
               .FirstOrDefault(row => row.Any(cell => cell.Contains("징수세액")))?
               .SkipWhile(cell => !cell.Contains("징수세액"))
               .Skip(1)
               .Take(3)
               .Select(cell => int.TryParse(cell.Trim().Replace(",", ""), out var value) ? value : 0)
               .Sum() ?? 0;

            // 디버깅용
            Console.WriteLine("------------");
            Console.WriteLine($"name: {name}");
            Console.WriteLine($"uid: {uid}");
            Console.WriteLine($"baseYear: {baseYear}");
            Console.WriteLine($"totalSum: {totalSum}");
            Console.WriteLine($"untaxedTotalSum: {untaxedTotalSum}");
            Console.WriteLine($"previousTaxPaid: {previousTaxPaid}");
            Console.WriteLine($"deductibleTax: {deductibleTax}");

            List<List<string>> secondTableData = PdfManager.ImportPdfToTable(filePath, 2, 3, 25);

            secondTableData = secondTableData
               .Select(row => row.Select(cell => Regex.Replace(cell, @"\s+", "")).ToList())
               .ToList();

            // 디버깅용
            foreach (var row in secondTableData)
            {
                Console.WriteLine(string.Join("|", row));
            }

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

            int preCalculatedSalary = Convert.ToInt32(totalSum + untaxedTotalSum - previousTaxPaid -
                (nationalPension + publicOfficialPension + soldierPension + privateSchoolPension + postalPension + healthInsurance + employmentInsurance));

            // 디버깅용
            Console.WriteLine("------------");
            Console.WriteLine($"nationalPension: {nationalPension}");
            Console.WriteLine($"publicOfficialPension: {publicOfficialPension}");
            Console.WriteLine($"soldierPension: {soldierPension}");
            Console.WriteLine($"privateSchoolPension: {privateSchoolPension}");
            Console.WriteLine($"postalPension: {postalPension}");
            Console.WriteLine($"healthInsurance: {healthInsurance}");
            Console.WriteLine($"employmentInsurance: {employmentInsurance}");
            Console.WriteLine($"preCalculatedSalary: {preCalculatedSalary}");

            return new PdfEmployeeData(name, uid, baseYear, preCalculatedSalary, deductibleTax);
        }
    }

    public static class PdfEmployeeDataExtractorFactory
    {
        public static IPdfEmployeeDataExtractor GetExtractor(int year)
        {
            switch (year)
            {
                case 2024:
                    return new PdfEmployeeDataExtractor2024();

                default:
                    return new PdfEmployeeDataExtractor2024();
            }
        }
    }

    // 위까지 DLL 분리

    internal class PdfService
    {
        public static List<PdfEmployeeData> ExtractPdfEmployeeDataFromFiles(List<string> filePaths)
        {
            List<PdfEmployeeData> employeeDataList = new List<PdfEmployeeData>();

            foreach (var filePath in filePaths)
            {
                try
                {
                    var extractor = PdfEmployeeDataExtractorFactory.GetExtractor(DateTime.Now.Year - 1);
                    PdfEmployeeData pdfEmployeeData = extractor.ExtractPdfEmployeeData(filePath);
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
