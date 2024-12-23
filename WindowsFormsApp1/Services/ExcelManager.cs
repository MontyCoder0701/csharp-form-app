using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using Newtonsoft.Json.Linq;

namespace WindowsFormsApp1
{
    internal class ExcelManager
    {
        public static void ExportJsonToExcel(string filePath, string jsonData)
        {
            JArray dataArray = JArray.Parse(jsonData);

            using (XLWorkbook workbook = new XLWorkbook())
            {
                IXLWorksheet worksheet = workbook.Worksheets.Add();

                var headers = dataArray.First?.ToObject<JObject>()?.Properties().Select(p => p.Name).ToList();
                if (headers == null || !headers.Any())
                {
                    throw new ArgumentException();
                }

                headers.Select((header, i) => worksheet.Cell(1, i + 1).Value = header).ToList();

                for (int rowIndex = 0; rowIndex < dataArray.Count; rowIndex++)
                {
                    JObject rowObj = (JObject)dataArray[rowIndex];
                    for (int colIndex = 0; colIndex < headers.Count; colIndex++)
                    {
                        string header = headers[colIndex];
                        worksheet.Cell(rowIndex + 2, colIndex + 1).Value = rowObj[header]?.ToString() ?? string.Empty;
                    }
                }

                worksheet.Range(1, 1, dataArray.Count + 1, headers.Count).CreateTable();

                workbook.SaveAs(filePath);
            }
        }

        public static string ImportExcelToJson(string filePath)
        {
            using (XLWorkbook workbook = new XLWorkbook(filePath))
            {
                IXLWorksheet worksheet = workbook.Worksheets.First();
                JArray dataArray = new JArray();
                List<IXLRow> rows = worksheet.RowsUsed().Skip(1).ToList();
                List<string> headers = worksheet.Row(1).CellsUsed().Select(c => c.GetValue<string>()).ToList();

                foreach (IXLRow row in rows)
                {
                    JObject obj = new JObject();
                    for (int i = 0; i < headers.Count; i++)
                    {
                        string cellValue = row.Cell(i + 1).Value.ToString();
                        obj[headers[i]] = cellValue;
                    }
                    dataArray.Add(obj);
                }

                return dataArray.ToString();
            }
        }
    }
}
