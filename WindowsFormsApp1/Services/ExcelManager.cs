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
                JObject headers = dataArray[0] as JObject;
                int col = 1;
                int row = 2;

                // TODO: 헤더셀 잠금처리 고려
                foreach (JProperty header in headers.Properties())
                {
                    worksheet.Cell(1, col++).Value = header.Name;
                }

                foreach (JToken item in dataArray)
                {
                    col = 1;
                    foreach (JToken value in item.Values())
                    {
                        IXLCell cell = worksheet.Cell(row, col++);
                        // TODO: 엑셀 셀 양식 확인
                        if (decimal.TryParse(value.ToString(), out decimal numericValue))
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
