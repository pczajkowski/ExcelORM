using System.Reflection;
using ClosedXML.Excel;
using ExcelORM.Attributes;

namespace ExcelORM;

public static class TypeExtensions
{
    // Borrowed from https://github.com/ClosedXML/ClosedXML/blob/develop/ClosedXML/Excel/XLCellValue.cs#L361
    public static object? ToObject(this XLCellValue value)
    {
        return value.Type switch
        {
            XLDataType.Blank => null,
            XLDataType.Boolean => value.GetBoolean(),
            XLDataType.Number => value.GetNumber(),
            XLDataType.Text => value.GetText(),
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
        var valueToSet = value.ToObject();
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