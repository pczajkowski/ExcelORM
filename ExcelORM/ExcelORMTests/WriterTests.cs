using ExcelORM;

namespace ExcelORMTests;

public class WriterTests
{
    private readonly Test[] arrayOfThree = 
    {
        new() { Name = "Bilbo", Surname = "Baggins", Job = "Eater"},
        new() { Name = "John", Job = "Policeman"},
        new() { Name = "Bruce", Surname = "Lee", Job = "Fighter"}
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
        var readArray = reader.Read<Test>(worksheetName).ToArray();
        Assert.Equal(3, readArray.Length);

        for (int i = 0; i < readArray.Length; i++)
            Assert.Equal(arrayOfThree[i], readArray[i]);

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
        writer.Write(arrayOfThree);
        writer.SaveAs(testFile);

        var reader = new ExcelReader(testFile);
        Assert.Equal(3, reader.Read<Test>().Count());

        writer.Write(listOfTwo, append: true);
        writer.SaveAs(testFile);
        
        reader = new ExcelReader(testFile);
        Assert.Equal(5, reader.Read<Test>().Count());
        File.Delete(testFile);
    }
    
    [Fact]
    public void WriteDifferentTypes()
    {
        var testFile = Path.GetRandomFileName();
        testFile = Path.ChangeExtension(testFile, "xlsx");

        var expected = new TestTypes
        {
            Date = DateTime.Now,
            TimeSpan = TimeSpan.MaxValue,
            Double = 2.33,
            Int = 1024,
            Text = "Test"
        };
        
        var list = new List<TestTypes>{ expected };
        
        var writer = new ExcelWriter(testFile);
        writer.Write(list);
        writer.SaveAs(testFile);

        var reader = new ExcelReader(testFile);
        var result = reader.Read<TestTypes>().ToList();
        Assert.Single(result);
        var first = result.First();
        Assert.Equal(expected.Date.ToString(), first.Date.ToString());
        Assert.Equal(expected.TimeSpan, first.TimeSpan);
        Assert.Equal(expected.Double, first.Double);
        Assert.Equal(expected.Int, first.Int);
        Assert.Equal(expected.Text, first.Text);

        File.Delete(testFile);
    }
}