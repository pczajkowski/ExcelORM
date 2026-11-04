using System.Reflection;
using ClosedXML.Excel;
using ExcelORM.Attributes;

namespace ExcelORM
{
    public class Mapping
    {
        public string? PropertyName { get; set; }
        public int? Position { get; set; }

        private static int? GetColumnPosition(IXLCells headerCells, PropertyInfo property, bool fallbackToPropertyName)
        {
            int? position = null;
            var columnAttributeObject = property.GetCustomAttributes(typeof(ColumnAttribute), false).FirstOrDefault();
            if (columnAttributeObject != null)
            {
                if (columnAttributeObject is ColumnAttribute { Names.Length: > 0 } columnAttribute)
                {
                    position = headerCells.FirstOrDefault(x => !x.Value.IsBlank && Array.Exists(columnAttribute.Names,
                            y => y.Equals(x.Value.ToString(), StringComparison.OrdinalIgnoreCase)))?.Address
                        .ColumnNumber;
                }
                
                if (position != null || !fallbackToPropertyName) return position;
            }
            
            return headerCells.FirstOrDefault(x =>
                    !x.Value.IsBlank &&
                    property.Name.Equals(x.Value.ToString(), StringComparison.OrdinalIgnoreCase))?.Address
                .ColumnNumber;
        }
        
        public static List<Mapping> MapProperties<T>(IXLCells? headerCells, bool fallbackToPropertyName = false) where T : class
        {
            var map = new List<Mapping>();
            if (headerCells == null || !headerCells.Any()) return map;

            var properties = typeof(T).GetProperties();
            foreach (var property in properties)
            {
                if (property.Skip()) continue;

                var position = GetColumnPosition(headerCells, property, fallbackToPropertyName);
                if (position == null) continue;
                
                map.Add(new Mapping { PropertyName = property.Name, Position = position });
            }

            return map;
        }
    }
}
