using System;
using System.Linq;
using ClosedXML.Excel;
using Newtonsoft.Json;

namespace ClosedXML.Report.Options.ConditionalFormatting
{
    public class ConditionalFormattingTag : OptionTag
    {
        private string Json => Parameters.ContainsKey("json") ? Parameters["json"] : null;

        public override void Execute(ProcessingContext context)
        {
            var ws = Range.Worksheet;

            // select a column
            var columnRange = ws.Range(context.Range.FirstCell().Address.RowNumber, this.Column,
                context.Range.LastCell().Address.RowNumber, this.Column);

            var data = JsonConvert.DeserializeObject<ConditionalFormattingModel>(Json);

            if (data.ColorScale != null)
            {
                columnRange.AddConditionalFormat().ColorScale()
                    .LowestValue(XLColor.FromHtml(data.ColorScale.LowestValue))
                    .HighestValue(XLColor.FromHtml(data.ColorScale.HighestValue));
            }

            columnRange.AddConditionalFormat().ColorScale()
                .LowestValue(XLColor.Teal)
                .HighestValue(XLColor.Orange);

            foreach (var rule in data.OperatorRules)
            {
                IXLStyle style;
                switch (rule.Operator)
                {
                    case XLCFOperator.Between:
                        style = columnRange.AddConditionalFormat()
                            .WhenBetween(rule.MinValue.Value, rule.MaxValue.Value);
                        break;
                    case XLCFOperator.GreaterThan:
                        style = columnRange.AddConditionalFormat().WhenGreaterThan(rule.Value.Value);
                        break;
                    case XLCFOperator.LessThan:
                        style = columnRange.AddConditionalFormat().WhenLessThan(rule.Value.Value);
                        break;
                    case XLCFOperator.EqualOrGreaterThan:
                        style = columnRange.AddConditionalFormat().WhenEqualOrGreaterThan(rule.Value.Value);
                        break;
                    case XLCFOperator.EqualOrLessThan:
                        style = columnRange.AddConditionalFormat().WhenEqualOrLessThan(rule.Value.Value);
                        break;
                    case XLCFOperator.NotBetween:
                        style = columnRange.AddConditionalFormat()
                            .WhenNotBetween(rule.MinValue.Value, rule.MaxValue.Value);
                        break;
                    case XLCFOperator.Equal:
                    case XLCFOperator.NotEqual:
                    case XLCFOperator.Contains:
                    case XLCFOperator.NotContains:
                    case XLCFOperator.StartsWith:
                    case XLCFOperator.EndsWith:
                    default:
                        throw new NotImplementedException(
                            $"The operator {rule.Operator.ToString()} is not currently not implemented.");
                }

                if (!string.IsNullOrEmpty(rule.Style.Background))
                    style.Fill.SetBackgroundColor(XLColor.FromHtml(rule.Style.Background));
                if (!string.IsNullOrEmpty(rule.Style.FontColor))
                    style.Font.SetFontColor(XLColor.FromHtml(rule.Style.FontColor));
                if (!string.IsNullOrEmpty(rule.Style.FontStyle))
                {
                    switch (rule.Style.FontStyle.ToLower())
                    {
                        case "bold":
                            style.Font.SetBold(true);
                            break;
                        case "italic":
                            style.Font.SetItalic(true);
                            break;
                        default:
                            throw new NotImplementedException(
                                $"Style {rule.Style.FontStyle} has not been implemented");
                    }
                }

                if (rule.Style.Size.HasValue)
                {
                    style.Font.SetFontSize(rule.Style.Size.Value);
                }
            }
        }

        public const string DefaultTagName = "ConditionalFormat";
        public const int DefaultTagPriority = 255;
    }
}
