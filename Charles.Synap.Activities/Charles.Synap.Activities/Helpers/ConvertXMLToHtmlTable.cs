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
