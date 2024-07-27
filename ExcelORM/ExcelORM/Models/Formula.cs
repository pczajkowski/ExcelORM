using ClosedXML.Excel;

namespace ExcelORM.Models
{
    public class Formula : SpecialBase
    {
        public object? Value { get; set; }
        public string? FormulaA1 { get; set; }
        public override void SetCellValue(IXLCell cell) => cell.FormulaA1 = FormulaA1;
        public override void GetValueFromCell(IXLCell cell)
        {
            Value = cell.Value.ToObject();
            if (cell.HasFormula) FormulaA1 = cell.FormulaA1;
        }
    }
}
