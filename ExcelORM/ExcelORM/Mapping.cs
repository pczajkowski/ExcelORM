using ClosedXML.Excel;

namespace ExcelORM
{
    public class Mapping
    {
        public string? PropertyName { get; set; }
        public int? Position { get; set; }

        public static List<Mapping>? MapProperties<T>(IXLCells? headerCells) where T : new()
        {
            if (headerCells == null || !headerCells.Any()) return null;

            var map = new List<Mapping>();
            var properties = typeof(T).GetProperties();
            foreach (var property in properties)
            {
                var position = property.GetCustomAttributes(typeof(ColumnAttribute), false).FirstOrDefault() is ColumnAttribute { Names.Length: > 0 } columnAttribute
                    ?
                    headerCells.FirstOrDefault(x => !x.Value.IsBlank && Array.Exists(columnAttribute.Names, y => y.Equals(x.Value.ToString(), StringComparison.InvariantCultureIgnoreCase)))?.Address.ColumnNumber
                    : headerCells.FirstOrDefault(x => !x.Value.IsBlank && x.Value.ToString().Equals(property.Name, StringComparison.InvariantCultureIgnoreCase))?.Address.ColumnNumber;

                if (position == null) continue;
                map.Add(new Mapping { PropertyName = property.Name, Position = position });
            }

            if (map.Count == properties.Length)
                return map;

            return null;
        }
    }
}
