using ClosedXML.Excel;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;

namespace ClosedXML.Report.Options.ConditionalFormatting
{
    public class ConditionalFormatting
    {
        public class StyleFormat
        {
            public string Background { get; set; }
            public string FontColor { get; set; }
            public string FontStyle { get; set; }
            public int? Size { get; set; }
        }

        [JsonConverter(typeof(StringEnumConverter))]
        public XLCFOperator Operator { get; set; }

        public int? MinValue { get; set; }
        public int? MaxValue { get; set; }
        public int? Value { get; set; }
        public StyleFormat Style { get; set; }
    }
}
