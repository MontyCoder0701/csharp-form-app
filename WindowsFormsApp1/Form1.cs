using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using ExcelLib;
using Newtonsoft.Json;
using PdfLib;
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
                    emplName: "오승진",
                    uidnum7: "8204181",
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

                    // DTO와 칼람명 Dictionary를 활용하여 저장할 JSON 형태로 변환합니다.
                    string jsonData = JsonConvert.SerializeObject(excelDtos);
                    jsonData = ExcelHeaderDictionary.Aggregate(jsonData, (current, kv) => current.Replace(kv.Key, kv.Value));

                    // DLL에서 호출
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
                    // DLL에서 호출
                    string jsonData = ExcelManager.ImportExcelToJson(filePath);

                    // JSON에서 칼람명 변환 위해 Dictionary를 사용하시면 됩니다.
                    jsonData = ExcelHeaderDictionary.Aggregate(jsonData, (current, kv) => current.Replace(kv.Value, kv.Key));

                    // JSON을 역직렬화하여 올바른 DTO 객체 목록으로 뽑아냅니다. (List<EmployeeExcelDto>)
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

        // 단일 PDF 파일 업로드 시 활용
        private void HandlePdfFileClick(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "PDF 파일 (*.pdf)|*.pdf";
                if (openFileDialog.ShowDialog() != DialogResult.OK)
                {
                    return;
                }

                try
                {
                    string filePath = openFileDialog.FileName;
                    PdfEmployeeData pdfEmployeeData = PdfService.ExtractPdfEmployeeDataFromFiles(new List<string> { filePath }).FirstOrDefault();

                    if (pdfEmployeeData == null)
                    {
                        MessageBox.Show("올바른 원천진수영수증이 아닙니다. 마스킹이 없는 작년 기준년도의 영수증을 올려주세요.");
                        return;
                    }

                    // 여기서부터는 메인에서 DB와 함께 처리하는 흐름입니다.
                    // 밑은 한 직원의 정보 업데이트 하는 예시입니다.
                    Employee updatingEmployee = employees.First();

                    if (!(updatingEmployee.EmplName == pdfEmployeeData.Name && updatingEmployee.Uidnum7 == pdfEmployeeData.Uidnum7))
                    {
                        MessageBox.Show("해당 직원에 대한 올바른 원천진수영수증이 아닙니다.");
                        return;
                    }

                    int validDeductibleTax = (updatingEmployee.DeductibleTaxBaseYear == DateTime.Now.Year - 2)
                        ? updatingEmployee.DeductibleTax ?? 0
                        : 0;

                    updatingEmployee.SalaryAmt = pdfEmployeeData.PreCalculatedSalary - validDeductibleTax;
                    updatingEmployee.SalaryBaseYear = pdfEmployeeData.BaseYear;

                    updatingEmployee.DeductibleTax = pdfEmployeeData.DeductibleTax;
                    updatingEmployee.DeductibleTaxBaseYear = pdfEmployeeData.BaseYear;

                    MessageBox.Show($"{pdfEmployeeData.Name}의 급여가 등록되었습니다.");

                }
                catch (Exception err)
                {
                    MessageBox.Show($"PDF 업로드 중 문제가 발생했습니다. {err.Message}");
                }
            }
        }

        // 여러 PDF 폴더 업로드 시 활용
        private void HandlePdfFolderClick(object sender, EventArgs e)
        {
            using (FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog())
            {
                if (folderBrowserDialog.ShowDialog() != DialogResult.OK)
                {
                    return;
                }

                try
                {
                    string selectedFolder = folderBrowserDialog.SelectedPath;
                    List<string> pdfFiles = Directory.GetFiles(selectedFolder, "*.pdf").ToList();

                    if (pdfFiles.Count == 0)
                    {
                        MessageBox.Show("선택한 폴더에 PDF 파일이 없습니다.");
                        return;
                    }

                    List<PdfEmployeeData> pdfEmployeeDataList = PdfService.ExtractPdfEmployeeDataFromFiles(pdfFiles);

                    if (pdfEmployeeDataList.Count == 0)
                    {
                        MessageBox.Show("선택한 폴더에 올바른 원천징수영수증 파일이 없습니다. 마스킹이 없는 작년 기준년도의 영수증을 올려주세요.");
                        return;
                    }

                    // 여기서부터는 메인에서 DB와 함께 처리하는 흐름입니다.
                    // 밑은 예시 흐름입니다.
                    // 저는 처음부터 중복 여부를 전부 제거하고 시작했습니다만, 편한 방법을 사용하셔도 됩니다.
                    List<PdfEmployeeData> uniquePdfDataList = pdfEmployeeDataList
                        .GroupBy(emp => new { emp.Name, emp.Uidnum7 })
                        .Where(g => g.Count() == 1)
                        .SelectMany(g => g)
                        .ToList();

                    List<Employee> uniqueEmployeeList = employees
                      .GroupBy(emp => new { emp.EmplNum, emp.Uidnum7 })
                      .Where(g => g.Count() == 1)
                      .SelectMany(g => g)
                      .ToList();

                    foreach (PdfEmployeeData pdfEmployeeData in uniquePdfDataList)
                    {
                        foreach (Employee employee in uniqueEmployeeList)
                        {
                            if (employee.EmplName == pdfEmployeeData.Name && employee.Uidnum7 == pdfEmployeeData.Uidnum7)
                            {
                                int validDeductibleTax = (employee.DeductibleTaxBaseYear == DateTime.Now.Year - 2)
                                    ? employee.DeductibleTax ?? 0
                                    : 0;

                                employee.SalaryAmt = pdfEmployeeData.PreCalculatedSalary - validDeductibleTax;
                                employee.SalaryBaseYear = pdfEmployeeData.BaseYear;

                                employee.DeductibleTax = pdfEmployeeData.DeductibleTax;
                                employee.DeductibleTaxBaseYear = pdfEmployeeData.BaseYear;

                                MessageBox.Show($"{pdfEmployeeData.Name}의 급여가 등록되었습니다.");
                            }
                        }
                    }
                }
                catch (Exception err)
                {
                    MessageBox.Show($"PDF 업로드 중 문제가 발생했습니다. {err.Message}");
                }
            }
        }
    }
}
