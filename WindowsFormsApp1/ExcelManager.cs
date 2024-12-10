using System.IO;
using ClosedXML.Excel;
using Newtonsoft.Json.Linq;

namespace WindowsFormsApp1
{
    internal class ExcelManager
    {
        public static void ExportJsonToExcel(string fileName, string filePath, string jsonData)
        {

            var dataArray = JArray.Parse(jsonData);

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Data");

                var headers = dataArray[0] as JObject;
                int col = 1;
                foreach (var header in headers.Properties())
                {
                    worksheet.Cell(1, col++).Value = header.Name;
                }

                int row = 2;
                foreach (var item in dataArray)
                {
                    col = 1;
                    foreach (var value in item.Values())
                    {
                        worksheet.Cell(row, col++).Value = value.ToString();
                    }
                    row++;
                }

                string savePath = Path.Combine(filePath, $"{fileName}.xlsx");
                workbook.SaveAs(savePath);
            }
        }
    }
}
