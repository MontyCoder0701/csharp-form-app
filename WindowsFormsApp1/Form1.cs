﻿using System;
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
                    uidnum7: "1234567",
                    healthInsuranceDate: "2024-12-01",
                    emplNum: "EMP001",
                    grade: "Manager",
                    department: "HR",
                    salaryBcode: 101,
                    salaryBname: "Bank A",
                    salaryAcctnum: "123-456-789",
                    allowanceBcode: 102,
                    allowanceBname: "Bank B",
                    allowanceAcctnum: "987-654-321",
                    salaryAmt: 5000,
                    salaryBaseYear: 2024,
                    deductibleTax: 500,
                    deductibleTaxBaseYear: 2024,
                    employmentDate: "2020-01-15",
                    isSafe: 1,
                    registdate: "2024-12-01",
                    registdateformat: "yyyy-MM-dd",
                    memo: "First employee record"
                ),
                new Employee(
                    emplSeq: 2,
                    projCode: "P002",
                    cid: "CUST002",
                    emplName: "이수경",
                    uidnum7: "7654321",
                    healthInsuranceDate: "2024-12-05",
                    emplNum: "EMP002",
                    grade: "Staff",
                    department: "Finance",
                    salaryBcode: 103,
                    salaryBname: "Bank C",
                    salaryAcctnum: "321-654-987",
                    allowanceBcode: 104,
                    allowanceBname: "Bank D",
                    allowanceAcctnum: "789-123-456",
                    salaryAmt: 4000,
                    salaryBaseYear: 2024,
                    deductibleTax: 400,
                    deductibleTaxBaseYear: 2024,
                    employmentDate: "2021-06-01",
                    isSafe: 0,
                    registdate: "2024-12-01",
                    registdateformat: "yyyy-MM-dd",
                    memo: "Second employee record"
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
                        string data = JsonConvert.SerializeObject(employees);

                        ExcelManager.ExportJsonToExcel(filePath, data);

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
