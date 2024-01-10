using System.Reflection;
using ClosedXML.Excel;

namespace ExcelORM;

public static class TypeExtensions
{
    public static void SetValue<T>(this T currentObject, PropertyInfo property, IXLCell cell)
    {
        object? valueToSet = property.PropertyType switch
        {
            not null when property.PropertyType == typeof(string) => cell.Value.ToString(),
            not null when property.PropertyType == typeof(DateTime?) => cell.Value.IsDateTime ? cell.Value.GetDateTime() : null,
            not null when property.PropertyType == typeof(TimeSpan?) => cell.Value.IsTimeSpan ? cell.Value.GetTimeSpan() : null,
            not null when property.PropertyType == typeof(double?) => cell.Value.IsNumber ? cell.Value.GetNumber() : null,
            not null when property.PropertyType == typeof(int?) => cell.Value.IsNumber ? (int?)cell.Value.GetNumber() : null,
            _ => throw new NotSupportedException($"{property.PropertyType} isn't supported!")
        };
               
        if (valueToSet != null)
            property.SetValue(currentObject, valueToSet); 
    }
}