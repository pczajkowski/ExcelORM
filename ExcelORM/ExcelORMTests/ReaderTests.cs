using ExcelORM;

namespace ExcelORMTests;

public class ReaderTests
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
}