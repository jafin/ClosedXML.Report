using System.Collections.Generic;

namespace ClosedXML.Report.Options.ConditionalFormatting
{
    public class ConditionalFormattingModel
    {
        public List<ConditionalFormatting> Rules { get; set; } = new List<ConditionalFormatting>();
    }
}
