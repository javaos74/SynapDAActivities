using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace Charles.Synap.Activities.Helpers
{
    public class ConvertXMLToHtmlTable
    {
        public ConvertXMLToHtmlTable() { }

        private static bool IsCoordinateAttribute(string attributeName)
        {
            string[] coordinateAttributes = { "top", "bottom", "left", "right", "page", "value", "id" }; // 'value' 속성도 제거 대상에 포함할지 고려
            return coordinateAttributes.Contains(attributeName);
        }

        public void ConvertXmlTablesToExcel(string xmlFilePath, string excelFilePath)
        {
            //ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            
            ExcelPackage.License.SetNonCommercialPersonal("Charles Kim");

            using (var package = new ExcelPackage())
            {
                XDocument doc = XDocument.Load(xmlFilePath);
                var tables = doc.Descendants("table");

                int tableIndex = 1;
                foreach (var table in tables)
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add($"tab{tableIndex}");
                    int row = 1;

                    var occupiedCells = new HashSet<(int Row, int Col)>();
                    var rows = table.Descendants("tr");
                    foreach (var rowElement in rows)
                    {
                        int col = 1;
                        
                        var cells = rowElement.Descendants("td").Concat(rowElement.Descendants("th"));
                        foreach (var cell in cells)
                        {

                            while (occupiedCells.Contains((row, col)))
                            {
                                col++;
                            }
                            // td/th 아래의 모든 p 태그 안의 span 값들을 줄바꿈으로 연결
                            string cellValue = string.Join("\n", cell.Descendants("p").Select(p => p.Descendants("span").Select(s => s.Value.Trim()).FirstOrDefault() ?? "").Where(s => !string.IsNullOrEmpty(s)));
                            worksheet.Cells[row, col].Value = cellValue;

                            int rowspan = cell.Attribute("rowspan") != null && int.TryParse(cell.Attribute("rowspan").Value, out int r) ? r : 1;
                            int colspan = cell.Attribute("colspan") != null && int.TryParse(cell.Attribute("colspan").Value, out int c) ? c : 1;

                            // 4. rowspan 또는 colspan이 1보다 클 경우에만 병합을 수행합니다.
                            if (rowspan > 1 || colspan > 1)
                            {
                                worksheet.Cells[row, col, row + rowspan - 1, col + colspan - 1].Merge = true;
#if DEBUG
                                System.Console.WriteLine($"병합: 시작({row},{col}) -> 끝({row + rowspan - 1},{col + colspan - 1})");
#endif
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
                    tableIndex++;
                }

                FileInfo excelFile = new FileInfo(excelFilePath);
                package.SaveAs(excelFile);
            }
        }
        private static XElement CleanElement(XElement element)
        {
            XElement cleanElement = new XElement(element.Name);
            foreach (var attribute in element.Attributes().Where(attr => !IsCoordinateAttribute(attr.Name.LocalName)))
            {
                cleanElement.Add(attribute);
            }
            foreach (var childNode in element.Nodes())
            {
                if (childNode is XElement childElement)
                {
                    cleanElement.Add(CleanElement(childElement));
                }
                else if (childNode is XText text)
                {
                    cleanElement.Add(text);
                }
                else if (childNode is XComment comment)
                {
                    cleanElement.Add(comment);
                }
            }
            return cleanElement;
        }

        public string ConvertTableToHtml(string xmlFilePath)
        {
            XDocument doc = XDocument.Load(xmlFilePath);
            StringBuilder extractedXml = new StringBuilder();

            var tableElements = doc.Descendants("table");
            foreach (var tableElement in tableElements)
            {
                // 새로운 table 요소 생성 및 속성 제거
                XElement cleanTableElement = new XElement("table");
                foreach (var attribute in tableElement.Attributes().Where(attr => !IsCoordinateAttribute(attr.Name.LocalName)))
                {
                    cleanTableElement.Add(attribute);
                }

                // 자식 요소 복사 및 좌표 속성 제거
                foreach (var childNode in tableElement.Nodes())
                {
                    if (childNode is XElement element)
                    {
                        cleanTableElement.Add(CleanElement(element));
                    }
                    else if (childNode is XText text)
                    {
                        cleanTableElement.Add(text);
                    }
                    else if (childNode is XComment comment)
                    {
                        cleanTableElement.Add(comment);
                    }
                }
                extractedXml.AppendLine(cleanTableElement.ToString());
            }

            return extractedXml.ToString();
        }
    }
}
