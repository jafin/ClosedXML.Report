using System.Collections.Generic;

namespace ClosedXML.Report.Options.ConditionalFormatting
{
    public class ConditionalFormattingModel
    {
        
        public List<ConditionalFormatting> OperatorRules { get; set; } = new List<ConditionalFormatting>();
        public ColorScale ColorScale { get; set; }

    }
}
