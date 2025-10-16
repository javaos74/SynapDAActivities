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
            Debugger.Launch();
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
                                    // (행, 열) 좌표를 키로 사용하여 이미 채워진 셀을 추적합니다.
                                    var occupiedCells = new HashSet<(int Row, int Col)>();
                                    var rows = table.Descendants("tr");
                                    foreach (var rowElement in rows)
                                    {
                                        int col = 1;
                                        var cells = rowElement.Descendants("td").Concat(rowElement.Descendants("th"));
                                        foreach (var cell in cells)
                                        {
                                            /*
                                            // td/th 아래의 모든 p 태그 안의 span 값들을 줄바꿈으로 연결
                                            */
                                            // 1. 현재 행에서 다음으로 데이터를 입력할 수 있는 '빈 셀'의 열(col)을 찾습니다.
                                            // 이전 row에서 rowspan으로 병합된 셀들을 건너뛰는 역할입니다.
                                            //while (worksheet.Cells[row, col].Merge)
                                            //{
                                            //    col++;
                                            //}

                                            // 현재 (rowIndex, colIndex)가 이미 rowspan으로 채워져 있다면
                                            // 비어있는 다음 열로 이동합니다.
                                            while (occupiedCells.Contains((row, col)))
                                            {
                                                col++;
                                            }
                                            // 2. 셀 값 추출
                                            string cellValue = string.Join("\n", cell.Descendants("p").Descendants("span").Select(s => s.Value.Trim()));
                                            worksheet.Cells[row, col].Value = cellValue;

                                            // 3. rowspan과 colspan 값을 파싱합니다. 속성이 없으면 기본값 1을 사용합니다.
                                            int rowspan = cell.Attribute("rowspan") != null && int.TryParse(cell.Attribute("rowspan").Value, out int r) ? r : 1;
                                            int colspan = cell.Attribute("colspan") != null && int.TryParse(cell.Attribute("colspan").Value, out int c) ? c : 1;

                                            // 4. rowspan 또는 colspan이 1보다 클 경우에만 병합을 수행합니다.
                                            if (rowspan > 1 || colspan > 1)
                                            {
                                                worksheet.Cells[row, col, row + rowspan - 1, col + colspan - 1].Merge = true;
                                            }
                                            // 병합으로 인해 차지하게 될 모든 셀을 'occupied'로 표시합니다.
                                            for (int rs = 0; rs < rowspan; rs++)
                                            {
                                                for (int cs = 0; cs < colspan; cs++)
                                                {
                                                    occupiedCells.Add((row + rs, col + cs));
                                                }
                                            }

                                            // 5. 현재 셀이 차지한 colspan 만큼 다음 셀의 시작 위치를 이동시킵니다.
                                            col += colspan;
                                        }
                                        row++;
                                    }
                                    tableCount++;
                                    tabIdx++;
                                    // 보기 좋게 열 너비를 자동 조정합니다.
                                    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
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
