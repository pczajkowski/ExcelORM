using ExcelORM;

namespace ExcelORMTests;

public class DynamicReaderTests
{
    private const string RegularFile = "testFiles/first.xlsx";
    
    [Fact]
    public void Read()
    {
        var reader = new ExcelDynamicReader(RegularFile);
        var results = reader.Read("Sheet 1").ToArray();
        Assert.NotEmpty(results);
    }
}