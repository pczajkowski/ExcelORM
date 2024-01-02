using ClosedXML.Excel;
using ExcelORM.Attributes;

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
                int? position;

                if (property.GetCustomAttributes(typeof(ColumnAttribute), false).FirstOrDefault() is ColumnAttribute { Names.Length: > 0 } attribute)
                    position = headerCells.FirstOrDefault(x => !x.Value.IsBlank && Array.Exists(attribute.Names,
                            y => y.Equals(x.Value.ToString(), StringComparison.InvariantCultureIgnoreCase)))?.Address
                        .ColumnNumber;
                else
                    position = headerCells.FirstOrDefault(x => !x.Value.IsBlank && property.Name.Equals(x.Value.ToString(), StringComparison.InvariantCultureIgnoreCase))?.Address.ColumnNumber;

                if (position == null) continue;
                map.Add(new Mapping { PropertyName = property.Name, Position = position });
            }

            return map.Count == properties.Length ? map : null;
        }
    }
}
