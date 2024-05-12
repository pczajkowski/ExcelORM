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
                if (property.Skip()) continue;

                var position = property.GetCustomAttributes(typeof(ColumnAttribute), false).FirstOrDefault() switch
                {
                    ColumnAttribute { Names.Length: > 0 } attribute => headerCells.FirstOrDefault(x => !x.Value.IsBlank && Array.Exists(attribute.Names,
                            y => y.Equals(x.Value.ToString(), StringComparison.InvariantCultureIgnoreCase)))?.Address
                        .ColumnNumber,
                    _ => headerCells.FirstOrDefault(x => !x.Value.IsBlank && property.Name.Equals(x.Value.ToString(), StringComparison.InvariantCultureIgnoreCase))?.Address.ColumnNumber
                };

                if (position == null) continue;
                map.Add(new Mapping { PropertyName = property.Name, Position = position });
            }

            return map.Count > 0 ? map : null;
        }
    }
}
