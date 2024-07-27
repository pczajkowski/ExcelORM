using ClosedXML.Excel;
using ExcelORM.Interfaces;

namespace ExcelORM.Models
{
    public class Formula : ISpecialProperty
    {
        public object? Value { get; set; }
        public string? FormulaA1 { get; set; }
        public void SetCellValue(IXLCell cell) => cell.FormulaA1 = FormulaA1;
    }
}
