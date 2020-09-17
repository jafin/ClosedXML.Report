using System.Linq;
using ClosedXML.Excel;

namespace ClosedXML.Report.Options
{
    public class ConditionalFormattingTag : OptionTag
    {
        //public XLSortOrder Order => Parameters.ContainsKey("desc") ? XLSortOrder.Descending : XLSortOrder.Ascending;

        //public int Num => Parameters.ContainsKey("num") ? Parameters["num"].AsInt(1) : int.MaxValue;

        public override void Execute(ProcessingContext context)
        {
            var fields = List.GetAll<ConditionalFormattingTag>().ToArray();
            foreach (var tag in fields)
            {
                //context.Range.AddConditionalFormat()
                context.Range.Columns(tag.Column, tag.Column);
                var ws = Range.Worksheet;

                // select a column
                var col = ws.Range(context.Range.FirstCell().Address.RowNumber, tag.Column,
                    context.Range.LastCell().Address.RowNumber, tag.Column);
                col.AddConditionalFormat().DataBar(XLColor.AliceBlue).LowestValue().HighestValue();
                //context.Range.SortColumns.Add(tag.Column, tag.Order);
            }

            context.Range.Sort();

            foreach (var tag in fields)
            {
                tag.Enabled = false;
            }
        }
    }
}
