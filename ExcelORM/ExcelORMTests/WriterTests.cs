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
        Assert.Equal(arrayOfThree.Length, reader.Read<Test>().Count());

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
            TimeSpan = TimeSpan.FromSeconds(360),
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
        Assert.NotNull(first.TimeSpan);
        Assert.Equal(expected.TimeSpan.Value.Minutes, first.TimeSpan.Value.Minutes);
        Assert.Equal(expected.Double, first.Double);
        Assert.Equal(expected.Int, first.Int);
        Assert.Equal(expected.Text, first.Text);

        File.Delete(testFile);
    }

    private readonly TestSkip[] arrayWithSkip =
    {
        new() {Text = "Lorem", Date = DateTime.Now.AddHours(1), Int = 1},
        new() {Text = "Ipsum", Date = null, Int = 2},
    };

    [Fact]
    public void WriteWithSkip()
    {
        var testFile = Path.GetRandomFileName();
        testFile = Path.ChangeExtension(testFile, "xlsx");

        const string worksheetName = "Test";
        var writer = new ExcelWriter(testFile);
        writer.Write(arrayWithSkip, worksheetName);
        writer.SaveAs(testFile);

        var reader = new ExcelReader(testFile);
        var readArray = reader.Read<TestSkip>(worksheetName).ToArray();
        Assert.Equal(arrayWithSkip.Length, readArray.Length);

        for (int i = 0; i < readArray.Length; i++)
        {
            Assert.NotEqual(arrayWithSkip[i].Text, readArray[i].Text);
            Assert.Equal(arrayWithSkip[i].Date.ToString(), readArray[i].Date.ToString());
            Assert.Equal(arrayWithSkip[i].Int, readArray[i].Int);
        }

        File.Delete(testFile);
    }

    private readonly TestSkipMiddle[] arrayWithSkipMiddle =
    {
        new() {Text = "Lorem", Date = DateTime.Now.AddHours(1), Int = 1},
        new() {Text = "Ipsum", Date = DateTime.Now.AddHours(2), Int = 2},
    };

    [Fact]
    public void WriteWithSkipMiddle()
    {
        var testFile = Path.GetRandomFileName();
        testFile = Path.ChangeExtension(testFile, "xlsx");

        const string worksheetName = "Test";
        var writer = new ExcelWriter(testFile);
        writer.Write(arrayWithSkipMiddle, worksheetName);
        writer.SaveAs(testFile);

        var reader = new ExcelReader(testFile);
        var readArray = reader.Read<TestSkipMiddle>(worksheetName).ToArray();
        Assert.Equal(arrayWithSkipMiddle.Length, readArray.Length);

        for (int i = 0; i < readArray.Length; i++)
        {
            Assert.Equal(arrayWithSkipMiddle[i].Text, readArray[i].Text);
            Assert.NotEqual(arrayWithSkipMiddle[i].Date.ToString(), readArray[i].Date.ToString());
            Assert.Equal(arrayWithSkipMiddle[i].Int, readArray[i].Int);
        }

        File.Delete(testFile);
    }
}