using ClosedXML.Excel;

namespace ExcelORM.Models
{
    public class Hyperlink : SpecialBase
    {
        public object? Value { get; set; }
        public XLHyperlink? Link { get; set; }
        public override void SetCellValue(IXLCell cell)
        {
            cell.Value = XLCellValue.FromObject(Value);
            cell.SetHyperlink(Link);
        }

        public override void GetValueFromCell(IXLCell cell)
        {
            Value = cell.Value.ToObject();
            if (cell.HasHyperlink) Link = cell.GetHyperlink();
        }
    }
}
