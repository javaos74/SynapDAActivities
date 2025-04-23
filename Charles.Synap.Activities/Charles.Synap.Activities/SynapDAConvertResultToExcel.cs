using System;
using System.Activities;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO.Abstractions;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Security.Cryptography;
using System.Security.Cryptography.Xml;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml.Linq;
using OfficeOpenXml;
using UiPath.Platform.ResourceHandling;
using UiPath.Platform.ResourceHandling.Internals;

namespace Charles.Synap.Activities
{
    public class SynapDAConvertResultToExcel : CodeActivity
    {
        public InArgument<IResource> ResultZip { get; set; }
        public InArgument<string> ResultExcelFile { get; set; }
        public OutArgument<int> TableCount { get; set; }
        public OutArgument<string> ErrorMessage { get; set; }

        private int tableCount; 

        public SynapDAConvertResultToExcel()
        {
            tableCount = 0;
        }
        private int ConvertXmlTablesToExcel(IResource zipFile, string excelFilePath)
        {
#if DEBUG
            //Debugger.Launch();
#endif
            //ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelPackage.License.SetNonCommercialPersonal("Charles Kim");
            ZipArchive archive = null;
            FileInfo excelFile = new FileInfo(excelFilePath);
            try
            {
                string tempRoot = Path.GetTempPath();
                string uniqueDirectoryName = Guid.NewGuid().ToString();
                string tempDirectoryPath = Path.Combine(tempRoot, uniqueDirectoryName);

                //using (FileStream fileStream = new FileStream(zipFilePath, FileMode.Open, FileAccess.Read, FileShare.Read))
                using (Stream fileStream = zipFile.GetReaderOrLocal().OpenStreamAsync().Result)
                {
                    using (archive = new ZipArchive(fileStream, ZipArchiveMode.Read))
                    {
                        // 압축 파일 내에서 확장자가 ".xml"인 파일들을 필터링합니다.
                        var xmlFiles = archive.Entries.Where(entry => entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) 
                                                                        && !entry.FullName.EndsWith("docinfo.xml", StringComparison.OrdinalIgnoreCase) )
                                                      .OrderBy(entry => entry.Name, StringComparer.OrdinalIgnoreCase);  
#if DEBUG
                        Console.WriteLine($"압축 파일 '{zipFile.FullName}'에서 XML 파일 추출 시작...");
#endif
                        int pageIdx = -1;
                        foreach (ZipArchiveEntry entry in xmlFiles)
                        {
                            string extractedFilePath = Path.Combine(tempDirectoryPath, entry.FullName);

                            // 디렉토리 구조를 유지하기 위해 필요한 하위 디렉토리를 생성합니다.
                            string directoryName = Path.GetDirectoryName(extractedFilePath);
                            if (!string.IsNullOrEmpty(directoryName) && !Directory.Exists(directoryName))
                            {
                                Directory.CreateDirectory(directoryName);
                            }
#if DEBUG
                            Console.WriteLine($"  '{entry.FullName}' 추출 중...");
#endif
                            entry.ExtractToFile(extractedFilePath, true); // true: 이미 파일이 존재하면 덮어쓰기
                            Match match = Regex.Match(entry.Name, @"_(\d{4})\.xml$");
                            if (match.Success)
                            {
                                // 그룹 1은 캡처된 네자리 숫자 문자열입니다.
                                string pageNumberString = match.Groups[1].Value;
                                // 문자열을 정수형으로 변환합니다.
                                if (!int.TryParse(pageNumberString, out pageIdx))
                                    pageIdx = -1;
                            }
                            using (var package = new ExcelPackage(excelFile))
                            {
                                XDocument doc = XDocument.Load(extractedFilePath);
                                var tables = doc.Descendants("table");
                                var tabIdx = 1;

                                foreach (var table in tables)
                                {
                                    if (package.Workbook.Worksheets.Any(ws => ws.Name == $"page{pageIdx}-tab{tabIdx}")) //exact matching
                                    {
                                        package.Workbook.Worksheets.Delete($"page{pageIdx}-tab{tabIdx}"); // 중복된 시트 이름이 있을 경우 삭제
                                    }
                                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add($"page{pageIdx}-tab{tabIdx}");
                                    int row = 1;

                                    var rows = table.Descendants("tr");
                                    foreach (var rowElement in rows)
                                    {
                                        int col = 1;
                                        var cells = rowElement.Descendants("td").Concat(rowElement.Descendants("th"));
                                        foreach (var cell in cells)
                                        {
                                            // td/th 아래의 모든 p 태그 안의 span 값들을 줄바꿈으로 연결
                                            //string cellValue = string.Join("\n", cell.Descendants("p").Select(p => p.Descendants("span").Select(s => s.Value.Trim()).FirstOrDefault() ?? "").Where(s => !string.IsNullOrEmpty(s)));
                                            string cellValue = string.Join("\n", cell.Descendants("p").Descendants("span").Select(s => s.Value));
                                            while (worksheet.Cells[row, col].Merge) // rowspan merge
                                            {
                                                col++;
                                            }
                                            worksheet.Cells[row, col].Value = cellValue ?? string.Empty;

                                            var colspanAttr = cell.Attribute("colspan");
                                            if (colspanAttr != null && int.TryParse(colspanAttr.Value, out int colspan))
                                            {
                                                worksheet.Cells[row, col, row, col + colspan - 1].Merge = true;
                                                col += colspan - 1;
                                            }
                                            var rowspanAttr = cell.Attribute("rowspan");
                                            if (rowspanAttr != null && int.TryParse(rowspanAttr.Value, out int rowspan))
                                            {
                                                worksheet.Cells[row, col, row + rowspan - 1, col].Merge = true;
                                            }
                                            col++;
                                        }
                                        row++;
                                    }
                                    tableCount++;
                                    tabIdx++;
                                }
#if DEBUG
                                Console.WriteLine($"  '{entry.FullName}' 추출 & 엑셀 table 변환 완료: '{extractedFilePath}'");
#endif
                                package.Save();
                            }
                            File.Delete(extractedFilePath); // 변환 후 XML 파일 삭제

                        }
                    }
                    Directory.Delete(tempDirectoryPath, true); // 변환 후 임시 디렉토리 삭제
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"오류 발생: {ex.Message}");
                tableCount = -1;
                if(archive != null)
                    archive.Dispose();  
            }
            
            return tableCount;
        }

        /*
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
                TableCount.Set(context, tableCount-1);
                ErrorMessage.Set(context, string.Empty);
            }
            else
            {
                ErrorMessage.Set(context, "Error in SynapDAConvertResultToExcel");
                TableCount.Set(context, 0);
            }
        } */

        protected override void Execute(CodeActivityContext context)
        {
            var zipfile = ResultZip.Get(context);
            var excelfile = ResultExcelFile.Get(context);

            if( ConvertXmlTablesToExcel(zipfile, excelfile) >= 0)
            {
                TableCount.Set(context, tableCount );
                ErrorMessage.Set(context, string.Empty);
            }
            else
            {
                ErrorMessage.Set(context, "Error in SynapDAConvertResultToExcel");
                TableCount.Set(context, 0); 
            }
        }
    }
}
