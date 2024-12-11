using System.IO;
using System.Linq;
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
                        var cell = worksheet.Cell(row, col++);

                        // TODO: Format 확인
                        if (decimal.TryParse(value.ToString(), out var numericValue))
                        {
                            cell.Value = numericValue;
                            cell.Style.NumberFormat.Format = "0";
                        }
                        else
                        {
                            cell.Value = value.ToString();
                        }
                    }
                    row++;
                }

                string savePath = Path.Combine(filePath, $"{fileName}.xlsx");
                workbook.SaveAs(savePath);
            }
        }

        public static string ImportExcelToJson(string filePath)
        {
            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheets.First();
                var rows = worksheet.RowsUsed().Skip(1);

                var headers = worksheet.Row(1).CellsUsed().Select(c => c.GetValue<string>()).ToList();
                var dataArray = new JArray();

                foreach (var row in rows)
                {
                    var obj = new JObject();
                    for (int i = 0; i < headers.Count; i++)
                    {
                        var cellValue = row.Cell(i + 1).Value.ToString();
                        obj[headers[i]] = cellValue;
                    }
                    dataArray.Add(obj);
                }

                return dataArray.ToString();
            }
        }
    }
}
