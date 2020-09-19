using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using ClosedXML.Excel;
using Newtonsoft.Json;

namespace ClosedXML.Report.Options
{
    public class FormattingRule
    {
        public string RuleType { get; set; }
        public int? From { get; set; }
        public int? To { get; set; }
        public int? Value { get; set; }
        public string Background { get; set; }
    }
    public class ConditionalFormattingModel
    {
        public ConditionalFormattingModel()
        {
        }

        public List<FormattingRule> Rules { get; set; } = new List<FormattingRule>();
    }

    public class ConditionalFormattingTag : OptionTag
    {
        public string Json => Parameters.ContainsKey("json") ? Parameters["json"] : null;

        public override void Execute(ProcessingContext context)
        {
            var fields = List.GetAll<ConditionalFormattingTag>().ToArray();
            foreach (var tag in fields)
            {
                var ws = Range.Worksheet;

                // select a column
                var col = ws.Range(context.Range.FirstCell().Address.RowNumber, tag.Column,
                    context.Range.LastCell().Address.RowNumber, tag.Column);
                //col.AddConditionalFormat().DataBar(GetColorOrDefault(Color)).LowestValue().HighestValue();
                // col.AddConditionalFormat().IconSet(XLIconSetStyle.ThreeTrafficLights1)
                //     .AddValue(XLCFIconSetOperator.EqualOrGreaterThan, 0, XLCFContentType.Number)
                //     .AddValue(XLCFIconSetOperator.EqualOrGreaterThan, 5000, XLCFContentType.Number)
                //     .AddValue(XLCFIconSetOperator.EqualOrGreaterThan, 10000, XLCFContentType.Number);

                var data = JsonConvert.DeserializeObject<ConditionalFormattingModel>(Json);
                foreach (var rule in data.Rules)
                {
                    switch (rule.RuleType)
                    {
                        case "between":
                            col.AddConditionalFormat().WhenBetween(rule.From.Value, rule.To.Value).Fill.SetBackgroundColor(XLColor.FromHtml(rule.Background));
                            break;
                        case "greaterThan":
                            col.AddConditionalFormat().WhenGreaterThan(rule.Value.Value).Fill.SetBackgroundColor(XLColor.FromHtml(rule.Background));
                            break;

                    }
                }

                // col.AddConditionalFormat().WhenBetween(0, 100).Fill.SetBackgroundColor(XLColor.FromHtml("#E2FEE2"));
                // col.AddConditionalFormat().WhenBetween(101, 10000).Fill.SetBackgroundColor(XLColor.FromHtml("#FFFFD4"));
                // col.AddConditionalFormat().WhenGreaterThan(10000).Fill.SetBackgroundColor(XLColor.FromHtml("#FB8383"));


                // col.AddConditionalFormat().ColorScale().LowestValue(XLColor.Green)
                //     .Midpoint(XLCFContentType.Number,50,XLColor.Yellow)
                //     .HighestValue(XLColor.Red);

            }
        }

        public const string DefaultTagName = "ConditionalFormat";
        public const int DefaultTagPriority = 0;
    }
}
