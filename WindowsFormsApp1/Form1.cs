using System;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
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
