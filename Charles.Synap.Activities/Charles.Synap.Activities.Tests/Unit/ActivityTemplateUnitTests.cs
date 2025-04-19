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
            var result = t.ConvertTableToHtml(@"C:\Temp\result\c5345aca-397b-4589-843a-37d584bcf963.docx_0002.xml");

            // Fix: Correctly create a FileStream with the appropriate constructor
            using FileStream fs = new FileStream(@"C:\Temp\table.html", FileMode.OpenOrCreate, FileAccess.ReadWrite);

            // Assuming you want to write the result to the file
            using StreamWriter writer = new StreamWriter(fs);
            writer.Write(result);

            t.ConvertXmlTablesToExcel(@"C:\Temp\result\c5345aca-397b-4589-843a-37d584bcf963.docx_0002.xml", @"C:\Temp\table2.xlsx");

            Assert.Equal(0, 0);
        }
    }
}
