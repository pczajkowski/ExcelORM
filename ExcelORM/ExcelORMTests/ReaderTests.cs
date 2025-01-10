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
    private const string WithFormulaFile = "testFiles/withFormula.xlsx";
    private const string BadDate = "testFiles/badDate.xlsx";
    
    [Fact]
    public void Read()
    {
        using var reader = new ExcelReader(RegularFile);
        var results = reader.Read<Test>().ToArray();
        Assert.NotEmpty(results);
    }
    
    [Fact]
    public void ReadHidden()
    {
        using var reader = new ExcelReader(HiddenFile);
        var results = reader.Read<Test>().ToArray();
        Assert.NotEmpty(results);

        using var readerHidden = new ExcelReader(HiddenFile) { SkipHidden = true };
        var resultsHidden = readerHidden.Read<Test>().ToArray();
        Assert.NotEmpty(resultsHidden);
        Assert.NotEqual(results.Length, resultsHidden.Length);
    }
    
    [Fact]
    public void ReadFiltered()
    {
        using var reader = new ExcelReader(FilteredFile);
        var results = reader.Read<Test>().ToArray();
        Assert.NotEmpty(results);

        using var readerFiltered = new ExcelReader(FilteredFile) { ObeyFilter = true };
        var resultsFiltered = readerFiltered.Read<Test>().ToArray();
        Assert.NotEmpty(resultsFiltered);
        Assert.NotEqual(results.Length, resultsFiltered.Length);
    }
    
    [Fact]
    public void ReadDifficult()
    {
        using var reader = new ExcelReader(DifficultFile);
        var results = reader.Read<Test>("Tab").ToArray();
        Assert.NotEmpty(results);

        var resultsWithTitle = reader.Read<Test>("WithTitle", startFrom: 2).ToArray();
        Assert.Equal(results.Length, resultsWithTitle.Length);

        var resultsBadHeader = reader.Read<Test>("BadHeader").ToArray();
        Assert.NotEmpty(resultsBadHeader);
        Assert.All(resultsBadHeader, x => Assert.Null(x.Name));
        Assert.All(resultsBadHeader, x => Assert.Null(x.Surname));
        Assert.All(resultsBadHeader, x => Assert.NotNull(x.Job));
    }
    
    [Fact]
    public void ReadMultipleSheets()
    {
        using var reader = new ExcelReader(MultipleSheetsFile);
        var results = reader.ReadAll<Test>().ToArray();
        Assert.NotEmpty(results);
        Assert.Equal(6, results.Length);
    }
    
    [Fact]
    public void ReadDifferentTypes()
    {
        using var reader = new ExcelReader(DifferentTypesFile);
        var results = reader.Read<TestTypes>().ToArray();
        Assert.NotEmpty(results);
    }

    [Fact]
    public void ReadDifferentTypesWithSkip()
    {
        using var reader = new ExcelReader(DifferentTypesFile);
        var results = reader.Read<TestSkip>().ToArray();
        Assert.NotEmpty(results);
        Assert.All(results, x => Assert.Null(x.Text));
        Assert.NotNull(results.FirstOrDefault(x => x.Date != null));
        Assert.NotNull(results.FirstOrDefault(x => x.Int != null));
    }

    [Fact]
    public void ReadDifferentTypesWithSkipMiddle()
    {
        using var reader = new ExcelReader(DifferentTypesFile);
        var results = reader.Read<TestSkipMiddle>().ToArray();
        Assert.NotEmpty(results);
        Assert.NotNull(results.FirstOrDefault(x => x.Text != null));
        Assert.All(results, x => Assert.Null(x.Date));
        Assert.NotNull(results.FirstOrDefault(x => x.Int != null));
    }

    [Fact]
    public void ReadWithFormula()
    {
        using var reader = new ExcelReader(WithFormulaFile);
        var results = reader.Read<TestWithFormula>().ToArray();
        Assert.NotEmpty(results);
    }
    
    [Fact]
    public void BadDateThrows()
    {
        using var reader = new ExcelReader(BadDate);
        Assert.Throws<ArgumentException>(() => reader.Read<TestTypes>().ToArray());
    }
}