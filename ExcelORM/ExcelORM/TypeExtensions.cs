using System.Reflection;
using ClosedXML.Excel;
using ExcelORM.Attributes;

namespace ExcelORM;

public static class TypeExtensions
{
    private static object? GetAdditionalTypeFromText(XLCellValue value, PropertyInfo? property = null)
    {
        if (property == null) return value.GetText();
        
        if (property.PropertyType == typeof(Guid?))
        {
            if (Guid.TryParse(value.GetText(), out var guid))
                return guid;

            return null;
        }

        if (Nullable.GetUnderlyingType(property.PropertyType) != null)
        {
            var genericType = property.PropertyType.GetGenericArguments().FirstOrDefault();
            if (genericType == null) return null;
            
            if (genericType.IsEnum)
                return Enum.TryParse(genericType, value.GetText(), true, out var enumValue)
                    ? enumValue : null;
        }
        
        return value.GetText(); 
    }
    
    // Borrowed from https://github.com/ClosedXML/ClosedXML/blob/develop/ClosedXML/Excel/XLCellValue.cs#L361
    public static object? ToObject(this XLCellValue value, PropertyInfo? property = null)
    {
        return value.Type switch
        {
            XLDataType.Blank => null,
            XLDataType.Boolean => value.GetBoolean(),
            XLDataType.Number => value.GetNumber(),
            XLDataType.Text => GetAdditionalTypeFromText(value, property),
            XLDataType.Error => value.GetError(),
            XLDataType.DateTime => value.GetDateTime(),
            XLDataType.TimeSpan => value.GetTimeSpan(),
            _ => throw new InvalidCastException()
        };
    }

    public static Type ValueType(this XLCellValue value)
    {
        return value.Type switch
        {
            XLDataType.Blank => typeof(string),
            XLDataType.Boolean => typeof(bool),
            XLDataType.Number => typeof(double?),
            XLDataType.Text => typeof(string),
            XLDataType.DateTime => typeof(DateTime?),
            XLDataType.TimeSpan => typeof(TimeSpan?),
            _ => throw new InvalidCastException()
        };
    }

    public static void SetPropertyValue<T>(this T currentObject, PropertyInfo property, XLCellValue value)
    {
        var valueToSet = value.ToObject(property);
        if (valueToSet == null) return;

        try
        {
            property.SetValue(currentObject, valueToSet);
        }
        catch
        {
            valueToSet = value.ToString();
            property.SetValue(currentObject, valueToSet);
        }
    }

    public static bool Skip(this PropertyInfo property) => property.GetCustomAttributes(typeof(SkipAttribute), false).FirstOrDefault() != null;
}