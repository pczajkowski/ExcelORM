using ExcelORM;

namespace ExcelORMTests;

public class ReaderTests
{
    private const string RegularFile = "testFiles/first.xlsx";
    private const string HiddenFile = "testFiles/hidden.xlsx";
    private const string FilteredFile = "testFiles/filtered.xlsx";
    private const string DifficultFile = "testFiles/columnsOnTheLeftHeaderNotFirstRow.xlsx";
    private const string MultipleSheetsFile = "testFiles/multipleSheets.xlsx";
    private const string DifferentTypesFile = "testFiles/differentTypes.xlsx";
    
    [Fact]
    public void Read()
    {
        var reader = new ExcelReader(RegularFile);
        var results = reader.Read<Test>().ToArray();
        Assert.NotEmpty(results);
    }
    
    [Fact]
    public void ReadHidden()
    {
        var reader = new ExcelReader(HiddenFile);
        var results = reader.Read<Test>().ToArray();
        Assert.NotEmpty(results);

        var readerHidden = new ExcelReader(HiddenFile) { SkipHidden = true };
        var resultsHidden = readerHidden.Read<Test>().ToArray();
        Assert.NotEmpty(resultsHidden);
        Assert.NotEqual(results.Length, resultsHidden.Length);
    }
    
    [Fact]
    public void ReadFiltered()
    {
        var reader = new ExcelReader(FilteredFile);
        var results = reader.Read<Test>().ToArray();
        Assert.NotEmpty(results);

        var readerFiltered = new ExcelReader(FilteredFile) { ObeyFilter = true };
        var resultsFiltered = readerFiltered.Read<Test>().ToArray();
        Assert.NotEmpty(resultsFiltered);
        Assert.NotEqual(results.Length, resultsFiltered.Length);
    }
    
    [Fact]
    public void ReadDifficult()
    {
        var reader = new ExcelReader(DifficultFile);
        var results = reader.Read<Test>("Tab").ToArray();
        Assert.NotEmpty(results);

        var resultsWithTitle = reader.Read<Test>("WithTitle", startFrom: 2).ToArray();
        Assert.Equal(results.Length, resultsWithTitle.Length);

        var resultsBadHeader = reader.Read<Test>("BadHeader").ToArray();
        Assert.Empty(resultsBadHeader);
    }
    
    [Fact]
    public void ReadMultipleSheets()
    {
        var reader = new ExcelReader(MultipleSheetsFile);
        var results = reader.Read<Test>().ToArray();
        Assert.NotEmpty(results);
        Assert.Equal(6, results.Length);
    }
    
    [Fact]
    public void ReadDifferentTypes()
    {
        var reader = new ExcelReader(DifferentTypesFile);
        var results = reader.Read<TestTypes>().ToArray();
        Assert.NotEmpty(results);
    }
}