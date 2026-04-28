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

        var pt = property.PropertyType;
        switch (pt)
        {
            case var _ when pt == typeof(Guid) || pt == typeof(Guid?):
                return HandleGuid(value, property);
            case var _ when pt == typeof(DateTime) || pt == typeof(DateTime?):
                DateTime.TryParse(value.GetText(), out var dateValue);
                return dateValue;
            case var _ when pt == typeof(DateOnly) || pt == typeof(DateOnly?):
                DateOnly.TryParse(value.GetText(), out var dateOnlyValue);
                return dateOnlyValue;
            case { IsEnum: true }:
            case var _ when Nullable.GetUnderlyingType(pt) is { IsEnum: true }:
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

    public static void SetPropertyValue<T>(this T currentObject, PropertyInfo property, XLCellValue value)
    {
        var valueToSet = value.ToObject(property);

        try
        {
            property.SetValue(currentObject, valueToSet);
        }
        catch (ArgumentException ex) // Catches issues like type mismatch or null for non-nullable
        {
            // If the property type is string, try to set it directly from the XLCellValue's string representation.
            if (property.PropertyType == typeof(string))
            {
                property.SetValue(currentObject, value.ToString());
            }
            else
            {
                // If it's not a string property, and the initial assignment failed,
                // re-throw with more context.
                throw new InvalidCastException($"Could not set property '{property.Name}' of type '{property.PropertyType.Name}' with value '{valueToSet}' (original XLCellValue: '{value}'). " +
                                               $"The value returned by ToObject was incompatible, and the property type is not string for fallback conversion. See inner exception for details.", ex);
            }
        }
        catch (Exception ex) // Catch any other unexpected exceptions from SetValue
        {
            throw new InvalidOperationException($"An unexpected error occurred while setting property '{property.Name}' of type '{property.PropertyType.Name}'. " +
                                                $"Value attempted to set: '{valueToSet}' (original XLCellValue: '{value}'). See inner exception for details.", ex);
        }
    }

    public static bool Skip(this PropertyInfo property) => property.GetCustomAttributes(typeof(SkipAttribute), false).FirstOrDefault() != null;
}