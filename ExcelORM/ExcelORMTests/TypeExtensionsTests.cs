using ClosedXML.Excel;
using ExcelORM;

namespace ExcelORMTests;

public class TypeExtensionsTests
{
    public DateTime? DateTimeProperty { get; set; }
    
    [Fact]
    public void ToObject_DateTimeAsString()
    {
        XLWorkbook wb = new();
        var ws = wb.Worksheets.Add("Sheet1");
        var cell = ws.Cell(1, 1).SetValue("7/27/2025");
        
        var propertyInfo = typeof(TypeExtensionsTests).GetProperty("DateTimeProperty");
        var readValue = cell.ToObject(propertyInfo);
        Assert.IsType<DateTime>(readValue);
    }
    
    public DateOnly? DateOnlyProperty { get; set; }
    
    [Fact]
    public void ToObject_DateOnlyAsString()
    {
        XLWorkbook wb = new();
        var ws = wb.Worksheets.Add("Sheet1");
        var cell = ws.Cell(1, 1).SetValue("7/27/2025");
        
        var propertyInfo = typeof(TypeExtensionsTests).GetProperty("DateOnlyProperty");
        var readValue = cell.ToObject(propertyInfo);
        Assert.IsType<DateOnly>(readValue);
    }
     
    public Guid? GuidProperty { get; set; }
    
    [Fact]
    public void ToObject_GuidAsString()
    {
        XLWorkbook wb = new();
        var ws = wb.Worksheets.Add("Sheet1");
        var cell = ws.Cell(1, 1).SetValue("00000000-0000-0000-0000-000000000001");
        
        var propertyInfo = typeof(TypeExtensionsTests).GetProperty("GuidProperty");
        var readValue = cell.ToObject(propertyInfo);
        Assert.IsType<Guid>(readValue);
        Assert.NotEqual(Guid.Empty, readValue);
    }
     
    public TestEnum? EnumProperty { get; set; }
    
    [Fact]
    public void ToObject_EnumAsString()
    {
        XLWorkbook wb = new();
        var ws = wb.Worksheets.Add("Sheet1");
        var cell = ws.Cell(1, 1).SetValue("Second");
        
        var propertyInfo = typeof(TypeExtensionsTests).GetProperty("EnumProperty");
        var readValue = cell.ToObject(propertyInfo);
        Assert.IsType<TestEnum>(readValue);
        Assert.NotEqual(TestEnum.First, readValue);
    }
    
    [Fact]
    public void GetTextQuotePrefix()
    {
        using XLWorkbook wb = new();
        var ws = wb.Worksheets.Add("Sheet1");
        var cell = ws.Cell(1, 1).SetValue("Test");
        cell.Style.SetIncludeQuotePrefix();
        
        var value = cell.GetTextQuotePrefix();
        Assert.Equal("'Test", value);
    }
}