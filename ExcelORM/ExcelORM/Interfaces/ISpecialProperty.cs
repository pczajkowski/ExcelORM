using ClosedXML.Excel;

namespace ExcelORM.Interfaces
{
    internal interface ISpecialProperty
    {
        public object? Value { get; set; }
        public void SetCellValue(IXLCell cell);
    }
}
