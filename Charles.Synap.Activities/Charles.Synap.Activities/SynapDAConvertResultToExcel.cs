using System;
using System.Activities;
using System.Collections.Generic;
using System.IO.Compression;
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
        private int ConvertXmlTablesToExcel(string zipFilePath, string excelFilePath)
        {
            //ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelPackage.License.SetNonCommercialPersonal("Charles Kim");
            ZipArchive archive = null;
            FileInfo excelFile = new FileInfo(excelFilePath);
            try
            {
                string tempRoot = Path.GetTempPath();
                string uniqueDirectoryName = Guid.NewGuid().ToString();
                string tempDirectoryPath = Path.Combine(tempRoot, uniqueDirectoryName);

                using (archive = ZipFile.OpenRead(zipFilePath))
                {
                    // 압축 파일 내에서 확장자가 ".xml"인 파일들을 필터링합니다.
                    var xmlFiles = archive.Entries.Where(entry => entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                                                  .OrderBy( entry => entry.Name, StringComparer.OrdinalIgnoreCase);
#if DEBUG
                    Console.WriteLine($"압축 파일 '{zipFilePath}'에서 XML 파일 추출 시작...");
#endif

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
                        using (var package = new ExcelPackage(excelFile))
                        {
                            XDocument doc = XDocument.Load(extractedFilePath);
                            var tables = doc.Descendants("table");

                            foreach (var table in tables)
                            {
                                if (package.Workbook.Worksheets.Any(ws => ws.Name == $"tab{tableIndex}")) //exact matching
                                {
                                    package.Workbook.Worksheets.Delete($"tab{tableIndex}"); // 중복된 시트 이름이 있을 경우 삭제
                                }
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
            catch (Exception ex)
            {
                Console.WriteLine($"오류 발생: {ex.Message}");
                tableIndex = -1;
                if(archive != null)
                    archive.Dispose();  
            }
            
            return tableIndex;
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
