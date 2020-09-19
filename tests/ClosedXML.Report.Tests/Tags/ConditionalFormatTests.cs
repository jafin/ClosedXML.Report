using System;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using ClosedXML.Report.Tests.TestModels;
using Xunit;
using Xunit.Abstractions;

namespace ClosedXML.Report.Tests.Tags
{
    public class ConditionalFormatTests : XlsxTemplateTestsBase, IDisposable
    {
        private IXLRange _rng;
        private XLWorkbook _workbook;
        private FileStream _stream;

        public ConditionalFormatTests(ITestOutputHelper output) : base(output)
        {
        }

        private void LoadTemplate(string fileTemplate)
        {
            var fileName = Path.Combine(TestConstants.TemplatesFolder, fileTemplate);
            _stream = File.Open(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            _workbook = new XLWorkbook(_stream);
            _rng = _workbook.Range("range1");
        }

        [Fact]
        public void BasicTest()
        {
            LoadTemplate("ConditionalFormat.xlsx");
            using (var db = new DbDemos())
            {
                var items = db.orders.ToList();
                using (var template = new XLTemplate(_stream))
                {
                    template.AddVariable("Orders", items);
                    template.Generate();
                    template.SaveAs(@"C:\temp\ConditionalFormatResult.xlsx");
                }
            }
        }

        public void Dispose()
        {
            _workbook?.Dispose();
            _stream?.Dispose();
        }
    }
}
