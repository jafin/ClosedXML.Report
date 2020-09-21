using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using ClosedXML.Report.Tests.TestModels;
using Xunit;
using Xunit.Abstractions;

namespace ClosedXML.Report.Tests.ConditionalFormat
{
    public class TemplateTests : XlsxTemplateTestsBase
    {
        public TemplateTests(ITestOutputHelper output) : base(output)
        {
        }

        [Fact]
        public void Blah()
        {
            var workBook = new XLWorkbook(@"C:\temp\ConditionalFormatOut2.xlsx");
            var sheet = workBook.Worksheet("Sheet1");
            var cell = sheet.Cell("J3");
            var newCell = sheet.Cell("J1");
            newCell.CopyFrom(cell);
            //workBook.Save();
        }



        [Theory,
         InlineData("ConditionalFormat.xlsx")]
        public void CanExecuteTemplate(string templateFile)
        {
            XlTemplateTest(templateFile, tpl =>
                {
                    tpl.AddVariable("Orders", WithOrders());
                    tpl.Generate();
                    tpl.SaveAs(@"C:\temp\ConditionalFormatOut.xlsx");
                },
                wb => CompareWithGauge(wb, templateFile));
        }

        private List<order> WithOrders()
        {
            using (var db = new DbDemos())
            {
                return db.orders.ToList();
            }
        }
    }
}
