using ExcelORM;

namespace ExcelORMTests;

public class ReaderWriterTests
{
    private const string regularFile = "testFiles/first.xlsx";
    private const string hiddenFile = "testFiles/hidden.xlsx";
    private const string filteredFile = "testFiles/filtered.xlsx";

    private class Test
    {
        [Column("First name" )]
        public string? Name { get; set; }

        [Column("Last name")]
        public string? Source { get; set; }

        [Column(new string[]{"Occupation", "Job"})]
        public string? Target { get; set; }
    }
    
    [Fact]
    public void Read()
    {
        var reader = new ExcelReader(regularFile);
        var results = reader.Read<Test>();
        Assert.NotNull(results);
        Assert.NotEmpty(results);
    }
    
    [Fact]
    public void ReadHidden()
    {
        var reader = new ExcelReader(hiddenFile);
        var results = reader.Read<Test>();
        Assert.NotNull(results);
        Assert.NotEmpty(results);

        var readerHidden = new ExcelReader(hiddenFile) { SkipHidden = true };
        var resultsHidden = readerHidden.Read<Test>();
        Assert.NotNull(resultsHidden);
        Assert.NotEmpty(resultsHidden);
        Assert.NotEqual(results.Count(), resultsHidden.Count());
    }
    
    [Fact]
    public void ReadFiltered()
    {
        var reader = new ExcelReader(filteredFile);
        var results = reader.Read<Test>();
        Assert.NotNull(results);
        Assert.NotEmpty(results);

        var readerFiltered = new ExcelReader(filteredFile) { ObeyFilter = true };
        var resultsFiltered = readerFiltered.Read<Test>();
        Assert.NotNull(resultsFiltered);
        Assert.NotEmpty(resultsFiltered);
        Assert.NotEqual(results.Count(), resultsFiltered.Count());
    }
}