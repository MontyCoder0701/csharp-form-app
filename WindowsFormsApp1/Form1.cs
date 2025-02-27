using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Newtonsoft.Json;
using WindowsFormsApp1.Models;
using WindowsFormsApp1.Services;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        private List<Employee> employees;
        private BindingSource employeesBindingSource;

        private static readonly Dictionary<string, string> ExcelHeaderDictionary = new Dictionary<string, string>()
        {
            { nameof(EmployeeExcelDto.EmplName), "직원 이름" },
            { nameof(EmployeeExcelDto.GetDisplayUidnum7), "주민등록번호 (앞 7자리)" },
            { nameof(EmployeeExcelDto.EmplNum), "직원 번호" },
            { nameof(EmployeeExcelDto.SalaryBname), "은행 이름" },
            { nameof(EmployeeExcelDto.SalaryAcctnum), "계좌 번호 (- 포함)" },
            { nameof(EmployeeExcelDto.SalaryAmt), "급여 금액 (원)" },
            { nameof(EmployeeExcelDto.SalaryBaseYear), "기준 연도" },
            { nameof(EmployeeExcelDto.EmploymentDate), "고용 날짜 (YYYY-MM-DD)" },
            { nameof(EmployeeExcelDto.GetDisplayIsSafe), "안전 여부 (O/X)" }
        };

        public Form1()
        {
            InitializeComponent();
            InitializeEmployees();
        }

        private void InitializeEmployees()
        {
            employees = new List<Employee>
            {
                new Employee(
                    emplSeq: 1,
                    cid: "GOTCHOO_1",
                    emplName: "이수정",
                    uidnum7: "9807012",
                    emplNum: "EMP001",
                    salaryBcode: 101,
                    salaryAcctnum: "123-456-789",
                    salaryAmt: 5000,
                    salaryBaseYear: 2024,
                    deductibleTax: 5000,
                    deductibleTaxBaseYear: 2023,
                    employmentDate: "2020-01-15",
                    isSafe: 1,
                    registdate: "2024-12-01",
                    registdateformat: "yyyy-MM-dd"
                ),
                new Employee(
                    emplSeq: 2,
                    cid: "GOTCHOO_1",
                    emplName: "이수경",
                    uidnum7: "0208202",
                    emplNum: "EMP002",
                    salaryBcode: 103,
                    salaryAcctnum: "321-654-987",
                    employmentDate: "2021-06-01",
                    isSafe: 0,
                    registdate: "2024-12-01",
                    registdateformat: "yyyy-MM-dd"
                )
            };

            employeesBindingSource = new BindingSource { DataSource = employees };
            dataGridView1.DataSource = employeesBindingSource;
        }

        private void HandleExportButtonClick(object sender, EventArgs e)
        {
            if (employees.Count == 0)
            {
                MessageBox.Show("다운로드 가능한 데이터가 없습니다.");
                return;
            }

            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "Excel 파일 (*.xlsx)|*.xlsx";
                saveFileDialog.FileName = $"직원정보_{DateTime.Now:yyyy-MM-dd}";

                if (saveFileDialog.ShowDialog() != DialogResult.OK)
                {
                    return;
                }

                try
                {
                    string filePath = saveFileDialog.FileName;

                    List<EmployeeExcelDto> excelDtos = employees.ConvertAll(emp => new EmployeeExcelDto
                    {
                        EmplName = emp.EmplName,
                        GetDisplayUidnum7 = emp.GetDisplayUidnum7,
                        EmplNum = emp.EmplNum,
                        SalaryBname = emp.SalaryBname,
                        SalaryAcctnum = emp.SalaryAcctnum,
                        SalaryAmt = emp.SalaryAmt,
                        SalaryBaseYear = emp.SalaryBaseYear,
                        EmploymentDate = emp.EmploymentDate,
                        GetDisplayIsSafe = emp.GetDisplayIsSafe
                    });

                    string jsonData = JsonConvert.SerializeObject(excelDtos);

                    jsonData = ExcelHeaderDictionary.Aggregate(jsonData, (current, kv) => current.Replace(kv.Key, kv.Value));

                    ExcelManager.ExportJsonToExcel(filePath, jsonData);

                    MessageBox.Show("파일이 저장되었습니다.");
                }
                catch (Exception err)
                {
                    MessageBox.Show($"엑셀 저장 중 문제가 발생했습니다. {err.Message}");
                }
            }
        }

        private void HandleUploadButtonClick(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show(
                "엑셀을 업로드하는 경우, 기존 데이터는 덮어씌워집니다. 진행하시겠습니까?",
                "경고",
                MessageBoxButtons.OKCancel,
                MessageBoxIcon.Warning
            );

            if (result != DialogResult.OK)
            {
                return;
            }

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel 파일 (*.xlsx)|*.xlsx";

                if (openFileDialog.ShowDialog() != DialogResult.OK)
                {
                    return;
                }

                try
                {
                    string filePath = openFileDialog.FileName;
                    string jsonData = ExcelManager.ImportExcelToJson(filePath);

                    jsonData = ExcelHeaderDictionary.Aggregate(jsonData, (current, kv) => current.Replace(kv.Value, kv.Key));

                    JsonSerializerSettings settings = new JsonSerializerSettings() { MissingMemberHandling = MissingMemberHandling.Error };
                    List<EmployeeExcelDto> newEmployeeDtos = JsonConvert.DeserializeObject<List<EmployeeExcelDto>>(jsonData, settings);

                    if (newEmployeeDtos == null || !newEmployeeDtos.Any())
                    {
                        throw new JsonSerializationException();
                    }

                    string validationError = newEmployeeDtos.Select(dto => dto.Validate()).FirstOrDefault(error => error != null);

                    if (validationError != null)
                    {
                        MessageBox.Show(validationError);
                        return;
                    }

                    foreach (EmployeeExcelDto dto in newEmployeeDtos)
                    {
                        Employee existingEmployee = employees.FirstOrDefault(emp =>
                            emp.EmplName == dto.EmplName && emp.Uidnum7 == dto.Uidnum7);

                        if (existingEmployee != null)
                        {
                            existingEmployee.EmplNum = dto.EmplNum;
                            existingEmployee.SalaryBcode = dto.SalaryBcode;
                            existingEmployee.SalaryAcctnum = dto.SalaryAcctnum;
                            existingEmployee.SalaryAmt = dto.SalaryAmt;
                            existingEmployee.SalaryBaseYear = dto.SalaryBaseYear;
                            existingEmployee.EmploymentDate = dto.EmploymentDate;
                            existingEmployee.IsSafe = dto.IsSafe;
                        }
                        else
                        {
                            Employee newEmployee = new Employee(
                                emplSeq: employees.Any() ? employees.Max(emp => emp.EmplSeq) + 1 : 1,
                                cid: "GOTCHOO_1",
                                emplName: dto.EmplName,
                                uidnum7: dto.Uidnum7,
                                emplNum: dto.EmplNum,
                                salaryBcode: dto.SalaryBcode,
                                salaryAcctnum: dto.SalaryAcctnum,
                                salaryAmt: dto.SalaryAmt,
                                salaryBaseYear: dto.SalaryBaseYear,
                                employmentDate: dto.EmploymentDate,
                                isSafe: dto.IsSafe,
                                registdate: DateTime.Now.ToString("yyyy-MM-dd"),
                                registdateformat: "yyyy-MM-dd"
                            );

                            employees.Add(newEmployee);
                        }
                    }

                    employeesBindingSource.ResetBindings(false);

                    MessageBox.Show("새로운 데이터가 업로드되었습니다.");
                }
                catch (JsonSerializationException)
                {
                    MessageBox.Show("제공된 기존 엑셀 틀이 잘못 바뀌었습니다. 다운받은 엑셀의 형태를 유지하여 다시 시도하세요.");
                }
                catch (Exception err)
                {
                    MessageBox.Show($"엑셀 업로드 중 문제가 발생했습니다. {err.Message}");
                }
            }
        }

        private void HandlePdfImportClick(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                // TODO: 여러 파일 선택 지원
                openFileDialog.Filter = "PDF 파일 (*.pdf)|*.pdf";

                if (openFileDialog.ShowDialog() != DialogResult.OK)
                {
                    return;
                }

                try
                {
                    // TODO: 2024 영수증도 확인 (더존, 세무사랑, 홈택스)
                    // 에러 처리 (중도입사자)
                    string filePath = openFileDialog.FileName;
                    List<List<string>> firstTableData = PdfManager.ImportPdfToTable(filePath, 1);

                    firstTableData = firstTableData
                       .Select(row => row.Select(cell => Regex.Replace(cell, @"\s+", "")).ToList())
                       .ToList();

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
                        .Last() ?? "";

                    if (name.Length == 0 || uid.Length == 0)
                    {
                        throw new Exception("잘못된 파일입니다. 올바른 원천징수영수증 원본을 업로드해주세요.");
                    }

                    if (name.Contains("*") || uid.Contains("*"))
                    {
                        throw new Exception("이름이나 주민번호가 가려져 직원 정보를 알 수 없습니다. 마스킹이 없는 파일을 업로드해주세요.");
                    }

                    string baseYear = firstTableData
                        .FirstOrDefault(row => row.Any(cell => cell.Contains("근무기간")))?
                        .SkipWhile(cell => !cell.Contains("근무기간"))
                        .Skip(1)
                        .FirstOrDefault()
                        .Substring(0, 4) ?? "";

                    if (string.IsNullOrWhiteSpace(baseYear) || !int.TryParse(baseYear, out int year))
                    {
                        throw new Exception("잘못된 파일입니다. 올바른 원천징수영수증 원본을 업로드해주세요.");
                    }

                    if (year != DateTime.Now.Year - 1)
                    {
                        throw new Exception("작년 기준년도의 원천징수영수증이 아닙니다. 올바른 파일을 업로드해주세요.");
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

                    Console.WriteLine($"name: {name}");
                    Console.WriteLine($"uid: {uid}");
                    Console.WriteLine($"baseYear: {baseYear}");
                    Console.WriteLine($"totalSum: {totalSum}");
                    Console.WriteLine($"untaxedTotalSum: {untaxedTotalSum}");
                    Console.WriteLine($"previousTaxPaid: {previousTaxPaid}");
                    Console.WriteLine($"deductibleTax : {deductibleTax}");

                    List<List<string>> secondTableData = PdfManager.ImportPdfToTable(filePath, 2, 3, 25);

                    secondTableData = secondTableData
                       .Select(row => row.Select(cell => Regex.Replace(cell, @"\s+", "")).ToList())
                       .ToList();

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

                    Console.WriteLine($"nationalPension: {nationalPension}");
                    Console.WriteLine($"publicOfficialPension: {publicOfficialPension}");
                    Console.WriteLine($"soldierPension: {soldierPension}");
                    Console.WriteLine($"privateSchoolPension: {privateSchoolPension}");
                    Console.WriteLine($"postalPension: {postalPension}");
                    Console.WriteLine($"healthInsurance: {healthInsurance}");
                    Console.WriteLine($"employmentInsurance: {employmentInsurance}");

                    // TODO: 급여액 = (근무처별소득명세 계 + 비과세소득 계) - 기납부세액 - [(국민연금보험료 or 공적연금보험료공제) + 건강보험료 + 고용보험료] - (전전년도 차감징수세액)
                    // 2025년이면 2024년 급여액, 2024년 차감징수세액으로 DB에 저장됌. 그다음해에는 올해 저장한 2024년 차감징수세액으로 업로드한 PDF 값들로 급여액을 다시 계산하여 업데이트 하고 무한 반복.
                    // 따라서 첫 업로드 해인 경우에는 2024 차감징수세액으로 급여액 계산 안함 (저저번년도 해 것을 쓰기 때문에)
                    decimal salary = totalSum + untaxedTotalSum - previousTaxPaid - (nationalPension + publicOfficialPension + soldierPension + privateSchoolPension + postalPension + healthInsurance + employmentInsurance);

                    MessageBox.Show($"이름: {name}, 기준년도: {baseYear}, 급여액: {salary}, 차감징수세액: {deductibleTax}");
                }
                catch (Exception err)
                {
                    MessageBox.Show($"PDF 업로드 중 문제가 발생했습니다. {err.Message}");
                }
            }
        }
    }
}
