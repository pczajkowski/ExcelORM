using System.Globalization;
using System.Reflection;
using ClosedXML.Excel;
using ExcelORM.Attributes;

namespace ExcelORM;

public static class TypeExtensions
{
    private static object? HandleGuid(XLCellValue value, PropertyInfo property)
    {
        if (Guid.TryParse(value.GetText(), out var guid))
            return guid;

        if (property.PropertyType == typeof(Guid?)) return null;
        return Guid.Empty; 
    }

    private static object? HandleEnum(XLCellValue value, PropertyInfo property, Type? nullableUnderlyingType)
    {
        if (nullableUnderlyingType != null)
        {
            return Enum.TryParse(nullableUnderlyingType, value.GetText(), true, out var enumNullableValue)
                ? enumNullableValue : null;
        }
        
        return Enum.TryParse(property.PropertyType, value.GetText(), true, out var enumValue)
            ? enumValue : Enum.GetValues(property.PropertyType).GetValue(0);
    } 
    
    private static object? GetAdditionalTypeFromText(XLCellValue value, PropertyInfo? property = null)
    {
        if (property == null) return value.GetText();

        switch (property.PropertyType)
        {
            case var guid when guid == typeof(Guid):
            case var guidNull when guidNull == typeof(Guid?):
                return HandleGuid(value, property);
            case var dateTime when dateTime == typeof(DateTime):
            case var dateTimeNull when dateTimeNull == typeof(DateTime?):
                return DateTime.TryParse(value.GetText(), out var dateValue) ? dateValue : default;
            case var dateOnly when dateOnly == typeof(DateOnly):
            case var dateOnlyNull when dateOnlyNull == typeof(DateOnly?):
                return DateOnly.TryParse(value.GetText(), out var dateOnlyValue) ? dateOnlyValue : default;
            case { IsEnum: true }:
            case var enumNull when Nullable.GetUnderlyingType(enumNull) is { IsEnum: true }:
                return HandleEnum(value, property, Nullable.GetUnderlyingType(property.PropertyType));
        }
        
        return value.GetText(); 
    }

    private static object? GetSpecificNumberType(XLCellValue value, PropertyInfo? property)
    {
        var rawNumber = value.GetNumber();
        if (property == null) return rawNumber;

        var targetType = Nullable.GetUnderlyingType(property.PropertyType) ?? property.PropertyType;
        if (!targetType.IsPrimitive && targetType != typeof(decimal)) return rawNumber;

        try
        {
            return Convert.ChangeType(rawNumber, targetType, CultureInfo.InvariantCulture);
        }
        catch (InvalidCastException)
        {
            return rawNumber;
        }
        catch (OverflowException)
        {
            if (Nullable.GetUnderlyingType(property.PropertyType) != null) return null;
            throw;
        }
    }
    
    // Borrowed from https://github.com/ClosedXML/ClosedXML/blob/develop/ClosedXML/Excel/XLCellValue.cs#L361
    public static object? ToObject(this XLCellValue value, PropertyInfo? property = null)
    {
        return value.Type switch
        {
            XLDataType.Blank => null,
            XLDataType.Boolean => value.GetBoolean(),
            XLDataType.Number => GetSpecificNumberType(value, property),
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