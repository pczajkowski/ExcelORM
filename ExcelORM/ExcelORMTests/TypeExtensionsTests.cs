using ClosedXML.Excel;
using ExcelORM;

namespace ExcelORMTests;

public class TypeExtensionsTests
{
    public DateTime? DateTimeProperty { get; set; }
    
    [Fact]
    public void ToObject_DateTimeAsString()
    {
        XLCellValue value = "7/27/2025";
        
        var propertyInfo = typeof(TypeExtensionsTests).GetProperty("DateTimeProperty");
        var readValue = value.ToObject(propertyInfo);
        Assert.IsType<DateTime>(readValue);
    }
    
    public DateOnly? DateOnlyProperty { get; set; }
    
    [Fact]
    public void ToObject_DateOnlyAsString()
    {
        XLCellValue value = "7/27/2025";
        
        var propertyInfo = typeof(TypeExtensionsTests).GetProperty("DateOnlyProperty");
        var readValue = value.ToObject(propertyInfo);
        Assert.IsType<DateOnly>(readValue);
    }
}