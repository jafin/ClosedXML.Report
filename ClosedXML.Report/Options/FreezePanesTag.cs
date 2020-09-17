using ClosedXML.Report.Utils;

namespace ClosedXML.Report.Options
{
    public class FreezePanesTag : OptionTag
    {
        private const string Workbook = "A1";
        private const string Worksheet = "A2";

        public const string DefaultTagName = "FreezePanes";
        public const byte DefaultTagPriority = 0;

        public class Props
        {
            public int? Rows { get; set; }
            public int? Columns { get; set; }
            public bool RowsAndColumnsHaveValue => Rows != null && Columns != null;
        }

        public override void Execute(ProcessingContext context)
        {
            var xlCell = Cell.GetXlCell(context.Range);
            var cellAddr = xlCell.Address.ToStringRelative(false);
            var ws = Range.Worksheet;

            var props = new Props();
            if (HasParameter("Rows"))
            {
                props.Rows = GetParameter("Rows").AsInt(0);
            }

            if (HasParameter("Columns"))
            {
                props.Columns = GetParameter("Columns").AsInt(0);
            }

            switch (cellAddr)
            {
                case Workbook:
                    //TODO;
                    break;
                case Worksheet:
                    if (props.RowsAndColumnsHaveValue)
                        ws.SheetView.Freeze(props.Rows.Value, props.Columns.Value);
                    else if (props.Rows.HasValue)
                    {
                        ws.SheetView.FreezeRows(props.Rows.Value);
                    }
                    else if (props.Columns.HasValue)
                    {
                        ws.SheetView.FreezeColumns(props.Columns.Value);
                    }

                    break;
            }
        }
    }
}
