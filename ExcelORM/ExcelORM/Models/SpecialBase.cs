using ClosedXML.Excel;

namespace ExcelORM.Models
{
    public class SpecialBase
    {
        public virtual void SetCellValue(IXLCell cell) { }
        public virtual void GetValueFromCell(IXLCell cell) { }
    }
}
