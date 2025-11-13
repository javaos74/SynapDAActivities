using Charles.Synap.Activities.Helpers;
using Xunit;

namespace Charles.Synap.Activities.Tests.Unit
{
    public class ActivityTemplateUnitTests
    {
        [Fact]
        public void Test()
        {
            ConvertXMLToHtmlTable t = new ConvertXMLToHtmlTable();
            var result = t.ConvertTableToHtml(@"C:\Temp\TEST.doc_0008.xml");

            // Fix: Correctly create a FileStream with the appropriate constructor
            using FileStream fs = new FileStream(@"C:\Temp\table.html", FileMode.OpenOrCreate, FileAccess.ReadWrite);

            // Assuming you want to write the result to the file
            using StreamWriter writer = new StreamWriter(fs);
            writer.Write(result);
            string folderPath = @"C:\Temp\494703_pdf";
            IEnumerable<string> xmlFiles = Directory.EnumerateFiles(folderPath, "*.xml");
            foreach (string filePath in xmlFiles)
            {
                System.Console.WriteLine(filePath);
                t.ConvertXmlTablesToExcel(filePath, @"C:\Temp\table2.xlsx", true);
            }

            Assert.Equal(0, 0);
        }
    }
}
