using System;
using System.Activities;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using OfficeOpenXml;

namespace Charles.Synap.Activities
{
    public class SynapDAConvertResultToExcel : AsyncCodeActivity
    {
        public InArgument<string> ResultZipFIle { get; set; }
        public InArgument<string> ResultExcelFile { get; set; }
        public OutArgument<int> TableCount { get; set; }
        public OutArgument<string> ErrorMessage { get; set; }

        private int tableIndex; 

        public SynapDAConvertResultToExcel()
        {
            tableIndex = 0;
        }
        private int ConvertXmlTablesToExcel(string xmlFilePath, string excelFilePath)
        {
            //ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelPackage.License.SetNonCommercialPersonal("Charles Kim");

            FileInfo excelFile = new FileInfo(excelFilePath);
            using (var package = new ExcelPackage( excelFile))
            {
                XDocument doc = XDocument.Load(xmlFilePath);
                var tables = doc.Descendants("table");

                foreach (var table in tables)
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add($"tab{tableIndex}");
                    int row = 1;

                    var rows = table.Descendants("tr");
                    foreach (var rowElement in rows)
                    {
                        int col = 1;
                        var cells = rowElement.Descendants("td").Concat(rowElement.Descendants("th"));
                        foreach (var cell in cells)
                        {
                            // td/th 아래의 모든 p 태그 안의 span 값들을 줄바꿈으로 연결
                            string cellValue = string.Join("\n", cell.Descendants("p").Select(p => p.Descendants("span").Select(s => s.Value.Trim()).FirstOrDefault() ?? "").Where(s => !string.IsNullOrEmpty(s)));
                            worksheet.Cells[row, col].Value = cellValue;

                            var colspanAttr = cell.Attribute("colspan");
                            if (colspanAttr != null && int.TryParse(colspanAttr.Value, out int colspan))
                            {
                                worksheet.Cells[row, col, row, col + colspan - 1].Merge = true;
                                col += colspan - 1;
                            }
                            col++;
                        }
                        row++;
                    }
                    tableIndex++;
                }
                package.Save();
                return tableIndex-1;
            }
        }

        protected override IAsyncResult BeginExecute(AsyncCodeActivityContext context, AsyncCallback callback, object state)
        {
            var zipfile = ResultZipFIle.Get(context);
            var excelfile = ResultExcelFile.Get(context);

            var task = new Task(_ => ConvertXmlTablesToExcel(zipfile, excelfile), state);
            task.Start();
            if (callback != null)
            {
                task.ContinueWith(s => callback(s));
                task.Wait();
            }
            return task;
        }

        protected override void EndExecute(AsyncCodeActivityContext context, IAsyncResult result)
        {
            var task = (Task)result;

            if (task.IsCompletedSuccessfully)
            {
                TableCount.Set(context, tableIndex-1);
                ErrorMessage.Set(context, string.Empty);
            }
            else
            {
                ErrorMessage.Set(context, "Error in SynapDARequest");
                TableCount.Set(context, 0);
            }
        }
    }
}
