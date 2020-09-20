using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using ClosedXML.Report.Options;
using ClosedXML.Report.Options.ConditionalFormatting;
using FluentAssertions;
using Xunit;

namespace ClosedXML.Report.Tests.ConditionalFormat
{
    public class TagTests : Tests.TagTests
    {
        [Fact]
        public void CanDeserializeTag()
        {
            var rng = _ws.Range("A1", "A5");
            _ws.Cell("A1").Value = 10.0;
            _ws.Cell("A2").Value = 20.0;
            _ws.Cell("A3").Value = 20.0;
            _ws.Cell("A4").Value = 100.0;

            var tag = CreateInRangeTag<ConditionalFormattingTag>(rng, rng.Cell("A5"));
            tag.Parameters = new Dictionary<string, string>
            {
                {
                    "json", @"{
           ""rules"": [
            {
            ""operator"": ""between"",
            ""minValue"": 0,
            ""maxValue"": 1000,
            ""background"": ""#E2FEE2""
        },
        {
            ""operator"": ""between"",
            ""minValue"": 1001,
            ""maxValue"": 10000,
            ""background"": ""#FFFFD4""
        },
        {
            ""operator"": ""greaterThan"",
            ""value"": 100001,
            ""background"": ""#FB8383""
            }
            ]
        }"
                }
            };
            tag.List = new TagsList(new TemplateErrors()) {tag};
            tag.Execute(new ProcessingContext(_ws.Range("A1", "A4"), new DataSource(new object[0])));

            var sheetConditionalFormats = _ws.RangeAddress.Worksheet.ConditionalFormats;

            // Assert sheet has conditional formatting
            sheetConditionalFormats.Count().Should().Be(3);
            var first = sheetConditionalFormats.First();

            // assert formatting rule has properties from json deserialize.
            first.Operator.Should().Be(XLCFOperator.Between);
            first.Values[1].Value.Should().Be("0");
            first.Values[2].Value.Should().Be("1000");
            first.Style.Fill.BackgroundColor.Color.Name.ToUpper().Should().Contain("E2FEE2");
        }
    }
}
