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

    private static object? GetAdditionalTypeFromText(IXLCell cell, PropertyInfo? property = null)
    {
        if (property == null) return cell.GetTextQuotePrefix();

        var pt = property.PropertyType;
        switch (pt)
        {
            case var _ when pt == typeof(Guid) || pt == typeof(Guid?):
                return HandleGuid(cell.Value, property);
            case var _ when pt == typeof(DateTime):
                if (DateTime.TryParse(cell.GetTextQuotePrefix(), out var dateValueNotNull)) return dateValueNotNull;
                break;
            case var _ when pt == typeof(DateTime?):
                if (DateTime.TryParse(cell.GetTextQuotePrefix(), out var dateValue)) return dateValue;
                return null;
            case var _ when pt == typeof(DateOnly):
                if (DateOnly.TryParse(cell.GetTextQuotePrefix(), out var dateOnlyValueNotNull)) return dateOnlyValueNotNull;
                break;
            case var _ when pt == typeof(DateOnly?):
                if (DateOnly.TryParse(cell.GetTextQuotePrefix(), out var dateOnlyValue)) return dateOnlyValue;
                return null;
            case { IsEnum: true }:
            case var _ when Nullable.GetUnderlyingType(pt) is { IsEnum: true }:
                return HandleEnum(cell.Value, property, Nullable.GetUnderlyingType(property.PropertyType));
        }

        return cell.GetTextQuotePrefix();
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
    public static object? ToObject(this IXLCell cell, PropertyInfo? property = null)
    {
        return cell.Value.Type switch
        {
            XLDataType.Blank => null,
            XLDataType.Boolean => cell.Value.GetBoolean(),
            XLDataType.Number => GetSpecificNumberType(cell.Value, property),
            XLDataType.Text => GetAdditionalTypeFromText(cell, property),
            XLDataType.Error => cell.Value.GetError(),
            XLDataType.DateTime => cell.Value.GetDateTime(),
            XLDataType.TimeSpan => cell.Value.GetTimeSpan(),
            _ => throw new InvalidCastException()
        };
    }

    public static void SetPropertyValue<T>(this T currentObject, PropertyInfo property, IXLCell cell)
    {
        var valueToSet = cell.ToObject(property);

        try
        {
            property.SetValue(currentObject, valueToSet);
        }
        catch (ArgumentException ex) // Catches issues like type mismatch or null for non-nullable
        {
            // If the property type is string, try to set it directly from the XLCellValue's string representation.
            if (property.PropertyType == typeof(string))
            {
                property.SetValue(currentObject, cell.GetTextQuotePrefix());
            }
            else
            {
                // If it's not a string property, and the initial assignment failed,
                // re-throw with more context.
                throw new InvalidCastException($"Could not set property '{property.Name}' of type '{property.PropertyType.Name}' with value '{valueToSet}' (original XLCellValue: '{cell.Value}'). " +
                                               $"The value returned by ToObject was incompatible, and the property type is not string for fallback conversion. See inner exception for details.", ex);
            }
        }
        catch (Exception ex) // Catch any other unexpected exceptions from SetValue
        {
            throw new InvalidOperationException($"An unexpected error occurred while setting property '{property.Name}' of type '{property.PropertyType.Name}'. " +
                                                $"Value attempted to set: '{valueToSet}' (original XLCellValue: '{cell.Value}'). See inner exception for details.", ex);
        }
    }

    public static bool Skip(this PropertyInfo property) => property.GetCustomAttributes(typeof(SkipAttribute), false).FirstOrDefault() != null;
    
    public static string GetTextQuotePrefix(this IXLCell? cell)
    {
        if (cell == null) return string.Empty;

        var text = cell.GetString();
        if (!cell.Value.IsText) return text;
        if (text.Length == 0) return string.Empty;
        
        return cell.Style.IncludeQuotePrefix ? $"'{text}" : text;
    }
}