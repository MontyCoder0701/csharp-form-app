using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using ClosedXML.Excel;
using Newtonsoft.Json;
using WindowsFormsApp1.Models;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        private List<Employee> employees;
        private BindingSource employeesBindingSource;

        private static readonly Dictionary<string, string> ExcelDictionary = new Dictionary<string, string>()
        {
            { "EmplName", "직원 이름" },
            { "DisplayUidnum7", "주민등록번호 (앞 7자리)" },
            { "EmplNum", "직원 번호" },
            { "SalaryBname", "은행 이름" },
            { "SalaryAcctnum", "계좌 번호" },
            { "SalaryAmt", "급여 금액" },
            { "SalaryBaseYear", "기준 연도" },
            { "EmploymentDate", "고용 날짜" },
            { "DisplayIsSafe", "안전 여부 (O/X)" }
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
                    salaryBname: "Bank A",
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
                    salaryBname: "Bank C",
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
                saveFileDialog.FileName = $"갖추_직원정보_{DateTime.Now:yyyy-MM-dd}";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        string filePath = saveFileDialog.FileName;

                        List<EmployeeExcelDto> excelDtos = employees.Select(emp => new EmployeeExcelDto
                        {
                            EmplName = emp.EmplName,
                            DisplayUidnum7 = emp.GetDisplayUidnum7,
                            EmplNum = emp.EmplNum,
                            SalaryBname = emp.SalaryBname,
                            SalaryAcctnum = emp.SalaryAcctnum,
                            SalaryAmt = emp.SalaryAmt,
                            SalaryBaseYear = emp.SalaryBaseYear,
                            EmploymentDate = emp.EmploymentDate,
                            DisplayIsSafe = emp.GetDisplayIsSafe
                        }).ToList();

                        string jsonData = JsonConvert.SerializeObject(excelDtos);

                        ExcelDictionary.ToList().ForEach(kv => jsonData = jsonData.Replace(kv.Key, kv.Value));

                        ExcelManager.ExportJsonToExcel(filePath, jsonData);

                        MessageBox.Show("파일이 저장되었습니다.");
                    }
                    catch (Exception err)
                    {
                        MessageBox.Show($"엑셀 저장 중 문제가 발생했습니다. {err.Message}");
                    }
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

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        string filePath = openFileDialog.FileName;
                        string jsonData = ExcelManager.ImportExcelToJson(filePath);

                        ExcelDictionary.ToList().ForEach(kv => jsonData = jsonData.Replace(kv.Value, kv.Key));

                        JsonSerializerSettings settings = new JsonSerializerSettings() { MissingMemberHandling = MissingMemberHandling.Error };
                        List<EmployeeExcelDto> newEmployeeDtos = JsonConvert.DeserializeObject<List<EmployeeExcelDto>>(jsonData, settings);

                        if (newEmployeeDtos == null || !newEmployeeDtos.Any())
                        {
                            throw new FormatException();
                        }

                        HashSet<(string EmplName, string Uidnum7)> existingEmployeeIds = employees
                            .Select(emp => (emp.EmplName, emp.Uidnum7))
                            .ToHashSet();

                        HashSet<(string EmplName, string Uidnum7)> newEmployeeIds = newEmployeeDtos
                            .Select(dto => (dto.EmplName, dto.Uidnum7))
                            .ToHashSet();

                        employees.RemoveAll(emp => !newEmployeeIds.Contains((emp.EmplName, emp.Uidnum7)));

                        foreach (EmployeeExcelDto dto in newEmployeeDtos)
                        {
                            Employee existingEmployee = employees.FirstOrDefault(emp =>
                                emp.EmplName == dto.EmplName && emp.Uidnum7 == dto.Uidnum7);

                            if (existingEmployee != null)
                            {
                                existingEmployee.EmplNum = dto.EmplNum;
                                existingEmployee.SalaryBname = dto.SalaryBname;
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
                                    salaryBcode: 0,
                                    salaryBname: dto.SalaryBname,
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
                        MessageBox.Show("입력된 데이터의 형식이 잘못되었습니다. 제공된 기존 엑셀 형식에 맞게 수정해주세요.");
                    }
                    catch (Exception err)
                    {
                        MessageBox.Show($"엑셀 업로드 중 문제가 발생했습니다. {err.Message}");
                    }
                }
            }
        }
    }
}
