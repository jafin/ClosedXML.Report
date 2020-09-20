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
            foreach (var tag in List.GetAll<ConditionalFormattingTag>().ToArray())
            {
                var ws = Range.Worksheet;

                // select a column
                var columnRange = ws.Range(context.Range.FirstCell().Address.RowNumber, tag.Column,
                    context.Range.LastCell().Address.RowNumber, tag.Column);

                var data = JsonConvert.DeserializeObject<ConditionalFormattingModel>(Json);
                foreach (var rule in data.Rules)
                {
                    switch (rule.Operator)
                    {
                        case XLCFOperator.Between:
                            columnRange.AddConditionalFormat().WhenBetween(rule.MinValue.Value, rule.MaxValue.Value).Fill
                                .SetBackgroundColor(XLColor.FromHtml(rule.Background));
                            break;
                        case XLCFOperator.GreaterThan:
                            columnRange.AddConditionalFormat().WhenGreaterThan(rule.Value.Value).Fill
                                .SetBackgroundColor(XLColor.FromHtml(rule.Background));
                            break;
                        case XLCFOperator.LessThan:
                            columnRange.AddConditionalFormat().WhenLessThan(rule.Value.Value).Fill
                                .SetBackgroundColor(XLColor.FromHtml(rule.Background));
                            break;
                        case XLCFOperator.EqualOrGreaterThan:
                            columnRange.AddConditionalFormat().WhenEqualOrGreaterThan(rule.Value.Value).Fill
                                .SetBackgroundColor(XLColor.FromHtml(rule.Background));
                            break;
                        case XLCFOperator.EqualOrLessThan:
                            columnRange.AddConditionalFormat().WhenEqualOrLessThan(rule.Value.Value).Fill
                                .SetBackgroundColor(XLColor.FromHtml(rule.Background));
                            break;
                        case XLCFOperator.NotBetween:
                            columnRange.AddConditionalFormat().WhenNotBetween(rule.MinValue.Value, rule.MaxValue.Value).Fill
                                .SetBackgroundColor(XLColor.FromHtml(rule.Background));
                            break;
                        // case XLCFOperator.Equal:
                        //     break;
                        // case XLCFOperator.NotEqual:
                        //     break;
                        // case XLCFOperator.Contains:
                        //     break;
                        // case XLCFOperator.NotContains:
                        //     break;
                        // case XLCFOperator.StartsWith:
                        //     break;
                        // case XLCFOperator.EndsWith:
                        //     break;
                        default:
                            throw new NotImplementedException(
                                $"The operator {rule.Operator.ToString()} is not currently not implemented.");
                    }
                }
            }
        }

        public const string DefaultTagName = "ConditionalFormat";
        public const int DefaultTagPriority = 0;
    }
}
