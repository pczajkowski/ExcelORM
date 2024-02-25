using ExcelORM;

namespace ExcelORMTests;

public class DynamicWriterTests
{
    private const string DifficultFile = "testFiles/dynamicDifficult.xlsx";

    [Fact]
    public void Write()
    {
        var testFile = Path.GetRandomFileName();
        testFile = Path.ChangeExtension(testFile, "xlsx");

        var reader = new ExcelDynamicReader(DifficultFile);
        var results = reader.Read().ToArray();
        Assert.NotEmpty(results);

        var writer = new ExcelDynamicWriter();
        writer.Write(results);
        writer.SaveAs(testFile);

        var savedReader = new ExcelDynamicReader(testFile);
        var savedResults = savedReader.Read().ToArray();
        Assert.NotEmpty(savedResults);
        Assert.True(results.First().SequenceEqual(savedResults.First()));
        Assert.True(results.Last().SequenceEqual(savedResults.Last()));

        File.Delete(testFile);
    }
}