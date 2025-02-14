using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Newtonsoft.Json;
using UglyToad.PdfPig;
using WindowsFormsApp1.Models;

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
                    salaryAmt: 4000,
                    salaryBaseYear: 2024,
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
                    // TODO: 2024 영수증도 확인
                    string filePath = openFileDialog.FileName;
                    List<List<string>> tableData = ImportPdfToTable(filePath, 2);

                    tableData = tableData
                       .Select(row => row.Select(cell => Regex.Replace(cell, @"\s+", "")).ToList())
                       .ToList();

                    foreach (var row in tableData)
                    {
                        Console.WriteLine(string.Join("|", row));
                    }

                    // TODO: 2024 영수증도 확인
                    var nameRow = tableData
                        .FirstOrDefault(row => row.Any(cell => cell.Contains("⑥") && row.Any(c => c.Contains("⑦"))))?
                        .SkipWhile(cell => !cell.Contains("⑥"))
                        .Skip(2)
                        .FirstOrDefault();

                    Console.WriteLine($"name: {nameRow ?? ""}");

                    var uid = tableData
                        .FirstOrDefault(row => row.Any(cell => cell.Contains("⑥") && row.Any(c => c.Contains("⑦"))))?
                        .Last();

                    Console.WriteLine($"uid: {uid ?? ""}");

                    // TODO: Regex 지양
                    string datePattern = @"\b\d{4}[-.]\d{2}[-.]\d{2}|\b\d{4}[-.]\d{2}";
                    var baseYear = tableData
                        .FirstOrDefault(row => row.Any(cell => cell.Contains("⑪") && cell.Contains("근무기간")))?
                        .SkipWhile(cell => !(cell.Contains("⑪") && cell.Contains("근무기간")))
                        .Skip(1)
                        .Select(cell => Regex.Match(cell, datePattern))
                        .FirstOrDefault(match => match.Success)?
                        .Value;

                    Console.WriteLine($"baseYear: {baseYear ?? ""}");

                    var totalSum = tableData
                        .FirstOrDefault(row => row.Any(cell => cell.Contains("16") && cell.Contains("계")))?
                       .SkipWhile(cell => !(cell.Contains("16") && cell.Contains("계")))
                       .Skip(1)
                       .FirstOrDefault();

                    Console.WriteLine($"totalSum: {totalSum ?? "0"}");

                    var untaxedTotalSum = tableData
                        .FirstOrDefault(row => row.Any(cell => cell.Contains("20") && cell.Contains("비과세소득")))?
                        .SkipWhile(cell => !(cell.Contains("20") && cell.Contains("비과세소득")))
                        .Skip(1)
                        .FirstOrDefault()?.Trim();

                    Console.WriteLine($"untaxedTotalSum: {(decimal.TryParse(untaxedTotalSum, out decimal untaxedSum) ? untaxedSum : 0)}");

                    var previousTaxPaid = tableData
                       .FirstOrDefault(row => row.Any(cell => cell.Contains("75") && cell.Contains("주(현)근무지")))?
                       .SkipWhile(cell => !(cell.Contains("75") && cell.Contains("주(현)근무지")))
                       .Skip(1)
                       .Take(3)
                       .Select(cell => decimal.TryParse(cell.Trim(), out decimal value) ? value : 0)
                       .Sum();

                    Console.WriteLine($"previousTaxPaid: {previousTaxPaid}");

                    MessageBox.Show("새로운 데이터가 업로드되었습니다.");
                }
                catch (Exception err)
                {
                    MessageBox.Show($"PDF 업로드 중 문제가 발생했습니다. {err.Message}");
                }
            }
        }

        // TODO: 별도 클래스에 분리
        List<List<string>> ImportPdfToTable(string pdfPath, int lastPage = 1, int rowThreshold = 5, int cellThreshold = 15)
        {
            List<List<(string text, double x, double y)>> tableRows = new List<List<(string, double, double)>>();

            using (PdfDocument document = PdfDocument.Open(pdfPath))
            {
                for (int pageNum = 1; pageNum <= lastPage; pageNum++)
                {
                    var page = document.GetPage(pageNum);
                    var words = page.GetWords()
                        .Select(word => (word.Text, word.BoundingBox.Left, word.BoundingBox.Bottom))
                        .ToList();

                    words.Sort((a, b) => b.Bottom.CompareTo(a.Bottom));

                    List<(string, double, double)> currentRow = new List<(string, double, double)>();
                    double lastY = double.MaxValue;

                    foreach (var (text, x, y) in words)
                    {
                        if (Math.Abs(y - lastY) > rowThreshold)
                        {
                            if (currentRow.Count > 0)
                            {
                                tableRows.Add(new List<(string, double, double)>(currentRow));
                            }
                            currentRow.Clear();
                        }
                        currentRow.Add((text, x, y));
                        lastY = y;
                    }

                    if (currentRow.Count > 0)
                    {
                        tableRows.Add(currentRow);
                    }
                }
            }

            List<List<string>> structuredTable = new List<List<string>>();
            foreach (var row in tableRows)
            {
                row.Sort((a, b) => a.x.CompareTo(b.x));

                List<string> mergedRow = new List<string>();
                string currentCell = row[0].text;
                double lastX = row[0].x;

                for (int i = 1; i < row.Count; i++)
                {
                    if (Math.Abs(row[i].x - lastX) < cellThreshold)
                    {
                        currentCell += " " + row[i].text;
                    }
                    else
                    {
                        mergedRow.Add(currentCell);
                        currentCell = row[i].text;
                    }
                    lastX = row[i].x;
                }
                mergedRow.Add(currentCell);
                structuredTable.Add(mergedRow);
            }

            return structuredTable;
        }
    }
}
