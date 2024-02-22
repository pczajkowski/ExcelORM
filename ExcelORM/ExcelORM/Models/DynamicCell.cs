using ClosedXML.Excel;

namespace ExcelORM.Models
{
    public record DynamicCell
    {
        public int Position { get; set; }
        public string? Header { get; set; }
        public Type? Type { get; set; }
        public object? Value { get; set; }

        public static List<DynamicCell>? MapHeader(IXLCells? headerCells)
        {
            if (headerCells == null || !headerCells.Any()) return null;

            var map = new List<DynamicCell>();
            foreach(var cell in headerCells)
            {
                var headerItem = new DynamicCell
                {
                    Position = cell.Address.ColumnNumber,
                    Header = cell.Value.GetText()
                };

                map.Add(headerItem);
            }

            return map;
        }
    }
}
