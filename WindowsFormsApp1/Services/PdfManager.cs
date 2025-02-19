using System;
using System.Collections.Generic;
using System.Linq;
using UglyToad.PdfPig;

namespace WindowsFormsApp1.Services
{
    internal class PdfManager
    {
        public static List<List<string>> ImportPdfToTable(string pdfPath, int pageNum = 1, int rowThreshold = 5, int cellThreshold = 30)
        {
            List<List<(string text, double x, double y)>> tableRows = new List<List<(string, double, double)>>();

            using (PdfDocument document = PdfDocument.Open(pdfPath))
            {

                var page = document.GetPage(pageNum);
                var words = page.GetWords()
                    .Select(word => (word.Text, word.BoundingBox.Left, word.BoundingBox.Bottom))
                    .ToList();

                words.Sort((a, b) => b.Bottom.CompareTo(a.Bottom));

                List<(string, double, double)> currentRow = new List<(string, double, double)>();
                double lastY = double.MaxValue;

                foreach (var (text, x, y) in words)
                {
                    if (Math.Abs(y - lastY) > rowThreshold)
                    {
                        if (currentRow.Count > 0)
                        {
                            tableRows.Add(new List<(string, double, double)>(currentRow));
                        }
                        currentRow.Clear();
                    }
                    currentRow.Add((text, x, y));
                    lastY = y;
                }

                if (currentRow.Count > 0)
                {
                    tableRows.Add(currentRow);
                }
            }

            List<List<string>> structuredTable = new List<List<string>>();

            foreach (var row in tableRows)
            {
                row.Sort((a, b) => a.x.CompareTo(b.x));

                List<string> mergedRow = new List<string>();
                string currentCell = row[0].text;
                double lastX = row[0].x;

                for (int i = 1; i < row.Count; i++)
                {
                    if (Math.Abs(row[i].x - lastX) < cellThreshold)
                    {
                        currentCell += " " + row[i].text;
                    }
                    else
                    {
                        mergedRow.Add(currentCell);
                        currentCell = row[i].text;
                    }
                    lastX = row[i].x;
                }
                mergedRow.Add(currentCell);
                structuredTable.Add(mergedRow);
            }

            return structuredTable;
        }
    }
}
