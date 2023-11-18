using ExcelORM;

namespace ExcelORMTests;

public class WriterTests
{
    private readonly Test[] arrayOfThree = 
    {
        new Test { Name = "Bilbo", Surname = "Baggins", Job = "Eater"},
        new Test { Name = "John", Surname = "McCain", Job = "Policeman"},
        new Test { Name = "Bruce", Surname = "Lee", Job = "Fighter"}
    };

    private readonly List<Test> listOfTwo = new()
    {
        new Test { Name = "Elon", Surname = "Musk", Job = "Comedian"},
        new Test { Name = "Donald", Surname = "Trump", Job = "Bankrupt"},
    };

    [Fact]
    public void WriteWithAppend()
    {
        var testFile = Path.GetRandomFileName();
        testFile = Path.ChangeExtension(testFile, "xlsx");

        const string worksheetName = "Test";
        var writer = new ExcelWriter(testFile);
        writer.Write(arrayOfThree, worksheetName);
        writer.SaveAs(testFile);

        var reader = new ExcelReader(testFile);
        Assert.Equal(3, reader.Read<Test>().Count());

        writer.Write(listOfTwo, worksheetName, true);
        writer.SaveAs(testFile);
        
        reader = new ExcelReader(testFile);
        Assert.Equal(5, reader.Read<Test>(worksheetName).Count());
        File.Delete(testFile);
    }
    
    [Fact]
    public void WriteWithAppendWithoutName()
    {
        var testFile = Path.GetRandomFileName();
        testFile = Path.ChangeExtension(testFile, "xlsx");

        var writer = new ExcelWriter(testFile);
        writer.Write(arrayOfThree, null);
        writer.SaveAs(testFile);

        var reader = new ExcelReader(testFile);
        Assert.Equal(3, reader.Read<Test>().Count());

        writer.Write(listOfTwo, null, true);
        writer.SaveAs(testFile);
        
        reader = new ExcelReader(testFile);
        Assert.Equal(5, reader.Read<Test>().Count());
        File.Delete(testFile);
    }
}