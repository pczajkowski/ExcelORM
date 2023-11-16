using ClosedXML.Excel;

namespace ExcelORM
{
    public class Mapping
    {
        public string? PropertyName { get; set; }
        public Type? PropertyType { get; set; }
        public int? Position { get; set; }

        public static List<Mapping>? MapProperties<T>(IXLCells? headerCells) where T : new()
        {
            if (headerCells == null || !headerCells.Any()) return null;

            var objectToRead = new T();
            var map = new List<Mapping>();
            var properties = objectToRead.GetType().GetProperties();
            foreach (var property in properties)
            {
                int? position = null;
                var columnAttribute = property.GetCustomAttributes(typeof(ColumnAttribute), false).FirstOrDefault() as ColumnAttribute;
                position = columnAttribute is { Names.Length: > 0 } ? headerCells.FirstOrDefault(x => !x.Value.IsBlank && Array.Exists(columnAttribute.Names, y => y.Equals(x.Value.ToString(), StringComparison.InvariantCultureIgnoreCase)))?.Address.ColumnNumber : headerCells.FirstOrDefault(x => !x.Value.IsBlank && x.Value.ToString().Equals(property.Name, StringComparison.InvariantCultureIgnoreCase))?.Address.ColumnNumber;

                if (position == null) continue;
                map.Add(new Mapping { PropertyName = property.Name, PropertyType = property.PropertyType, Position = position });
            }

            if (map.Count == properties.Length)
                return map;

            return null;
        }
    }
}
