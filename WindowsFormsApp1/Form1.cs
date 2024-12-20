using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Newtonsoft.Json;
using WindowsFormsApp1.Models;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        private List<Employee> employees;
        private BindingSource employeesBindingSource;

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
                    projCode: "P001",
                    cid: "CUST001",
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
                    projCode: "P002",
                    cid: "CUST002",
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

                        // 테이블 데이터 넘기듯 넘기면 됌
                        List<string> headers = new List<string>
                        {
                            "직원 이름", "주민등록번호", "직원 번호", "은행 코드",
                            "은행 이름", "계좌 번호", "급여 금액", "기준 연도",
                            "고용 날짜", "안전 여부"
                        };

                        List<List<dynamic>> data = employees.Select(emp => new List<dynamic>
                        {
                            emp.EmplName,
                            emp.GetDisplayUidnum7,
                            emp.EmplNum,
                            emp.SalaryBcode,
                            emp.SalaryBname,
                            emp.SalaryAcctnum,
                            emp.SalaryAmt,
                            emp.SalaryBaseYear,
                            emp.EmploymentDate,
                            emp.GetDisplayIsSafe,
                        }).ToList();

                        ExcelManager.ExportDataToExcel(filePath, headers, data);

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
                        string data = ExcelManager.ImportExcelToJson(filePath);

                        JsonSerializerSettings settings = new JsonSerializerSettings() { MissingMemberHandling = MissingMemberHandling.Error };
                        List<Employee> importedEmployees = JsonConvert.DeserializeObject<List<Employee>>(data, settings);

                        if (importedEmployees == null || importedEmployees.Any(employee => employee == null))
                        {
                            throw new JsonException();
                        }

                        employees.Clear();
                        employees.AddRange(importedEmployees);
                        employeesBindingSource.ResetBindings(false);

                        MessageBox.Show("새로운 데이터가 업로드되었습니다.");
                    }
                    catch (JsonException)
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
