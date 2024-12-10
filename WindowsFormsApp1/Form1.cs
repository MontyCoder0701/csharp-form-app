using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            InitializeTransactions();
        }

        private void InitializeTransactions()
        {
            List<BankTransaction> transactions = new List<BankTransaction>
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

            dataGridView1.DataSource = transactions;
        }

        private void HandleExportButtonClick(object sender, EventArgs e)
        {
            string data = @"[
                { 'Name': 'Alice', 'Age': 25, 'City': 'New York' },
                { 'Name': 'Bob', 'Age': 30, 'City': 'Los Angeles' },
                { 'Name': 'Charlie', 'Age': 22, 'City': 'Chicago' }
            ]";

            string fileName = "엑셀파일임";
            string filePath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            ExcelManager.ExportJsonToExcel(fileName, filePath, data);
            MessageBox.Show("파일이 저장되었습니다.");
        }
    }
}
