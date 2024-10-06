using ClosedXML.Excel;
using ExcelORM;

namespace ExcelORMTests;

public class DynamicWriterTests
{
    private const string DifficultFile = "testFiles/dynamicDifficult.xlsx";
    private const string MultipleSheetsFile = "testFiles/multipleSheets.xlsx";

    [Fact]
    public void Write()
    {
        var testFile = Path.GetRandomFileName();
        testFile = Path.ChangeExtension(testFile, "xlsx");

        using var reader = new ExcelDynamicReader(DifficultFile);
        var results = reader.Read().ToArray();
        Assert.NotEmpty(results);

        using var writer = new ExcelDynamicWriter();
        writer.Write(results);
        writer.SaveAs(testFile);

        using var savedReader = new ExcelDynamicReader(testFile);
        var savedResults = savedReader.Read().ToArray();
        Assert.NotEmpty(savedResults);
        Assert.True(results.First().SequenceEqual(savedResults.First()));
        Assert.True(results.Last().SequenceEqual(savedResults.Last()));

        File.Delete(testFile);
    }

    [Fact]
    public void WriteAll()
    {
        var testFile = Path.GetRandomFileName();
        testFile = Path.ChangeExtension(testFile, "xlsx");

        using var reader = new ExcelDynamicReader(MultipleSheetsFile);
        var results = reader.ReadAll().ToArray();
        Assert.NotEmpty(results);
        
        using var writer = new ExcelDynamicWriter();
        writer.WriteAll(results);
        writer.SaveAs(testFile);

        using var savedReader = new ExcelDynamicReader(testFile);
        var savedResults = savedReader.ReadAll().ToArray();
        Assert.NotEmpty(savedResults);
        Assert.Equal(results.First().Name, savedResults.First().Name);
        Assert.Equal(results.First().Cells?.Count(), savedResults.First().Cells?.Count());
        Assert.Equal(results.Last().Name, savedResults.Last().Name);
        Assert.Equal(results.Last().Cells?.Count(), savedResults.Last().Cells?.Count());

        File.Delete(testFile);
    }

    [Fact]
    public void WriteReadInMemory()
    {
        using var readWorkbook = new XLWorkbook(DifficultFile);
        using var reader = new ExcelDynamicReader(readWorkbook);
        var results = reader.Read().ToArray();
        Assert.NotEmpty(results);

        using var writeWorkbook = new XLWorkbook();
        using var writer = new ExcelDynamicWriter(writeWorkbook);
        writer.Write(results);

        using var savedReader = new ExcelDynamicReader(writeWorkbook);
        var savedResults = savedReader.Read().ToArray();
        Assert.NotEmpty(savedResults);
        Assert.True(results.First().SequenceEqual(savedResults.First()));
        Assert.True(results.Last().SequenceEqual(savedResults.Last()));
    }
}