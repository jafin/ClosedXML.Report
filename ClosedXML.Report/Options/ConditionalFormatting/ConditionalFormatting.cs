using ClosedXML.Excel;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;

namespace ClosedXML.Report.Options.ConditionalFormatting
{
    public class ConditionalFormatting
    {
        [JsonConverter(typeof(StringEnumConverter))]
        public XLCFOperator Operator { get; set; }

        public int? MinValue { get; set; }
        public int? MaxValue { get; set; }
        public int? Value { get; set; }
        public string Background { get; set; }
    }
}
