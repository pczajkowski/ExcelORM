using ExcelORM;

namespace ExcelORMTests;

public class DynamicReaderTests
{
    private const string RegularFile = "testFiles/first.xlsx";
    private const string DifferentTypesFile = "testFiles/differentTypes.xlsx";

    [Fact]
    public void Read()
    {
        var reader = new ExcelDynamicReader(RegularFile);
        var results = reader.Read("Sheet 1").ToArray();
        Assert.NotEmpty(results);
    }

    [Fact]
    public void ReadDifferentTypes()
    {
        var reader = new ExcelDynamicReader(DifferentTypesFile);
        var results = reader.Read("Sheet1").ToArray();
        Assert.NotEmpty(results);

        var first = results.First();
        Assert.Equal(typeof(string), first[0].Type);
        Assert.Equal(typeof(DateTime?), first[1].Type);
        Assert.Equal(typeof(TimeSpan?), first[2].Type);
        Assert.Equal(typeof(double?), first[3].Type);
        Assert.Equal(typeof(double?), first[4].Type);
    }
}