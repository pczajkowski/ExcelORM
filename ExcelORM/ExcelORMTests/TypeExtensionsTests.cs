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
     
    public Guid? GuidProperty { get; set; }
    
    [Fact]
    public void ToObject_GuidAsString()
    {
        XLCellValue value = "00000000-0000-0000-0000-000000000001";
        
        var propertyInfo = typeof(TypeExtensionsTests).GetProperty("GuidProperty");
        var readValue = value.ToObject(propertyInfo);
        Assert.IsType<Guid>(readValue);
        Assert.NotEqual(Guid.Empty, readValue);
    }
}