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
            using (var document = PdfDocument.Open(pdfPath))
            {
                var words = document.GetPage(pageNum).GetWords()
                    .Select(w => new { w.Text, X = w.BoundingBox.Left, Y = w.BoundingBox.Bottom })
                    .OrderByDescending(w => w.Y)
                    .ToList();

                var tableRows = new List<List<(string text, double x)>>();
                var currentRow = new List<(string, double)>();
                double lastY = double.MaxValue;

                foreach (var word in words)
                {
                    if (Math.Abs(word.Y - lastY) > rowThreshold && currentRow.Count > 0)
                    {
                        tableRows.Add(new List<(string, double)>(currentRow));
                        currentRow.Clear();
                    }
                    currentRow.Add((word.Text, word.X));
                    lastY = word.Y;
                }

                if (currentRow.Count > 0)
                {
                    tableRows.Add(currentRow);
                }

                var structuredTable = new List<List<string>>();

                foreach (var row in tableRows)
                {
                    row.Sort((a, b) => a.x.CompareTo(b.x));

                    var mergedRow = new List<string>();
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
}
