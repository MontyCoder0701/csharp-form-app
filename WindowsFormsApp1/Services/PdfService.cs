using System;
using System.Collections.Generic;
using PdfLib;

namespace WindowsFormsApp1.Services
{
    internal class PdfService
    {
        public static List<PdfEmployeeData> ExtractPdfEmployeeDataFromFiles(List<string> filePaths)
        {
            List<PdfEmployeeData> employeeDataList = new List<PdfEmployeeData>();

            foreach (string filePath in filePaths)
            {
                try
                {
                    IPdfEmployeeDataExtractor extractor = PdfEmployeeDataExtractorFactory.GetExtractor(DateTime.Now.Year - 1);
                    PdfEmployeeData pdfEmployeeData = extractor.ExtractPdfEmployeeData(filePath);
                    employeeDataList.Add(pdfEmployeeData);
                }
                catch
                {
                    continue;
                }
            }

            return employeeDataList;
        }
    }
}
