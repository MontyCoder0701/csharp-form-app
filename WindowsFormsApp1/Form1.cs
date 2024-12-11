using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using Newtonsoft.Json;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        private List<BankTransaction> transactions;
        private BindingSource transactionsBindingSource;

        public Form1()
        {
            InitializeComponent();
            InitializeTransactions();
        }

        private void InitializeTransactions()
        {
            transactions = new List<BankTransaction>
            {
                new BankTransaction(
                    tranSeq: 1,
                    acctSeq: 1001,
                    balance100: 50000,
                    bcode: 1234,
                    cid: "CUST001",
                    date: "2024-12-10",
                    description: "Transaction 1 Description",
                    order: 1,
                    pay100: 2000,
                    save100: 3000,
                    time: "10:30 AM",
                    trankind: "Debit",
                    transferSeq: 2001,
                    where: "ATM1"
                ),
                new BankTransaction(
                    tranSeq: 2,
                    acctSeq: 1002,
                    balance100: 100000,
                    bcode: 5678,
                    cid: "CUST002",
                    date: "2024-12-09",
                    description: "Transaction 2 Description",
                    order: 2,
                    pay100: 5000,
                    save100: 7000,
                    time: "11:15 AM",
                    trankind: "Credit",
                    transferSeq: 2002,
                    where: "Branch1"
                ),
            };

            transactionsBindingSource = new BindingSource { DataSource = transactions };
            dataGridView1.DataSource = transactionsBindingSource;
        }

        private void HandleExportButtonClick(object sender, EventArgs e)
        {
            // TODO: 경로 선택 창 추가
            string data = JsonConvert.SerializeObject(transactions);
            string fileName = "갖추_은행_거래내역";
            string filePath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            ExcelManager.ExportJsonToExcel(fileName, filePath, data);
            // TODO: 실패 에러처리 및 중복파일 검사
            MessageBox.Show("파일이 저장되었습니다.");
        }

        private void HandleUploadButtonClick(object sender, EventArgs e)
        {
            // TODO: 경로 선택 창 추가
            string fileName = "갖추_은행_거래내역";
            string filePath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string savePath = Path.Combine(filePath, $"{fileName}.xlsx");


            // TODO: 실패 에러처리
            string data = ExcelManager.ImportExcelToJson(savePath);
            var importedTransactions = JsonConvert.DeserializeObject<List<BankTransaction>>(data);

            transactions.Clear();
            transactions.AddRange(importedTransactions);
            transactionsBindingSource.ResetBindings();

            MessageBox.Show("새로운 데이터가 업로드되었습니다.");
        }
    }
}
