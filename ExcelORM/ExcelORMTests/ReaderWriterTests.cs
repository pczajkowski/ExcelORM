using ExcelORM;

namespace ExcelORMTests;

public class ReaderWriterTests
{
    private const string RegularFile = "testFiles/first.xlsx";
    private const string HiddenFile = "testFiles/hidden.xlsx";
    private const string FilteredFile = "testFiles/filtered.xlsx";

    private class Test
    {
        [Column("First name" )]
        public string? Name { get; set; }

        [Column("Last name")]
        public string? Surname { get; set; }

        [Column(new[]{"Occupation", "Job"})]
        public string? Job { get; set; }
    }
    
    [Fact]
    public void Read()
    {
        var reader = new ExcelReader(RegularFile);
        var results = reader.Read<Test>();
        Assert.NotNull(results);
        Assert.NotEmpty(results);
    }
    
    [Fact]
    public void ReadHidden()
    {
        var reader = new ExcelReader(HiddenFile);
        var results = reader.Read<Test>();
        Assert.NotNull(results);
        Assert.NotEmpty(results);

        var readerHidden = new ExcelReader(HiddenFile) { SkipHidden = true };
        var resultsHidden = readerHidden.Read<Test>();
        Assert.NotNull(resultsHidden);
        Assert.NotEmpty(resultsHidden);
        Assert.NotEqual(results.Count(), resultsHidden.Count());
    }
    
    [Fact]
    public void ReadFiltered()
    {
        var reader = new ExcelReader(FilteredFile);
        var results = reader.Read<Test>();
        Assert.NotNull(results);
        Assert.NotEmpty(results);

        var readerFiltered = new ExcelReader(FilteredFile) { ObeyFilter = true };
        var resultsFiltered = readerFiltered.Read<Test>();
        Assert.NotNull(resultsFiltered);
        Assert.NotEmpty(resultsFiltered);
        Assert.NotEqual(results.Count(), resultsFiltered.Count());
    }
}