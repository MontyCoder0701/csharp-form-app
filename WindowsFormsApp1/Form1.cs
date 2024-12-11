using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

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
            // TODO: 스키마 기반 임시로 만든 거래내역 클래스
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
            // TODO: 데이터 없는 경우 버튼 disable 등 다른 처리 고려 가능
            if (transactions.Count == 0)
            {
                MessageBox.Show("다운로드 가능한 데이터가 없습니다.");
                return;
            }


            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "Excel 파일 (*.xlsx)|*.xlsx";
                saveFileDialog.FileName = $"갖추_은행_거래내역_{DateTime.Now:yyyy-MM-dd}";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    // TODO: 에러핸들링 메세지 수정
                    try
                    {
                        string filePath = saveFileDialog.FileName;
                        string data = JsonConvert.SerializeObject(transactions);

                        // TODO: 번역 매핑, 날짜 형태, 소수점 형태 등 포매팅 추가
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
            // TODO: 데이터 덮어씌울지 결정 및 경고 멘트 등 수정
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
                    // TODO: 경우 별 에러핸들링 메세지 수정
                    try
                    {
                        string filePath = openFileDialog.FileName;
                        string data = ExcelManager.ImportExcelToJson(filePath);
                        JArray jsonArray = JArray.Parse(data);

                        if (jsonArray.Any(item => item.Type == JTokenType.Object && !item.HasValues))
                        {
                            throw new JsonException();
                        }

                        // TODO: 번역 매핑, 날짜 형태, 소수점 형태 등 포매팅 변환
                        List<BankTransaction> importedTransactions = jsonArray.ToObject<List<BankTransaction>>();

                        transactions.Clear();
                        transactions.AddRange(importedTransactions);
                        transactionsBindingSource.ResetBindings(false);

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
