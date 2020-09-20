using System.Collections.Generic;
using System.Linq;
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

        [Theory,
         InlineData("ConditionalFormat.xlsx")]
        public void CanExecuteTemplate(string templateFile)
        {
            XlTemplateTest(templateFile, tpl => tpl.AddVariable("Orders", WithOrders()),
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
