using ExcelORM;
using ExcelORM.Models;
using ClosedXML.Excel;

namespace ExcelORMTests;

public class WriterTests
{
    private static readonly Test[] arrayOfThree = 
    {
        new() { Name = "Bilbo", Surname = "Baggins", Job = "Eater"},
        new() { Name = "John", Job = "Policeman"},
        new() { Name = "Bruce", Surname = "Lee", Job = "Fighter"}
    };

    private static readonly List<Test> listOfTwo = new()
    {
        new() { Name = "Elon", Surname = "Musk", Job = "Comedian"},
        new() { Name = "Donald", Surname = "Trump", Job = "Bankrupt"},
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

    private const string ForAppend = "testFiles/forAppend.xlsx";

    [Fact]
    public void WriteWithAppendExisting()
    {
        var testFile = Path.GetRandomFileName();
        testFile = Path.ChangeExtension(testFile, "xlsx");
        File.Copy(ForAppend, testFile);

        uint headerRowIndex = 3;
        var writer = new ExcelWriter(testFile);
        writer.Write(arrayOfThree, append: true, headerRowIndex: headerRowIndex);
        writer.SaveAs(testFile);

        var reader = new ExcelReader(testFile);
        var readArray = reader.Read<Test>(startFrom: headerRowIndex).ToArray();
        Assert.Equal(6, readArray.Length);

        for (int i = 0; i < arrayOfThree.Length; i++)
            Assert.Equal(arrayOfThree[i], readArray[i+3]);

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

    private static readonly TestSkip[] arrayWithSkip =
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

    private static readonly TestSkipMiddle[] arrayWithSkipMiddle =
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

    private static readonly TestWithFormula[] arrayWithFormulas =
    {
        new() { Name = "Bilbo", Surname = "Baggins", Job = "Eater", FullName = new Formula{ FormulaA1 = "B2&C2" } },
        new() { Name = "John", Job = "Policeman", FullName = new Formula{ FormulaA1 = "B3&C3" } },
        new() { Name = "Bruce", Surname = "Lee", Job = "Fighter", FullName = new Formula{ FormulaA1 = "B4&C4" } },
    };

    [Fact]
    public void WriteWithFormula()
    {
        var testFile = Path.GetRandomFileName();
        testFile = Path.ChangeExtension(testFile, "xlsx");

        var writer = new ExcelWriter(testFile);
        writer.Write(arrayWithFormulas);
        writer.SaveAs(testFile);

        var reader = new ExcelReader(testFile);
        var readArray = reader.Read<TestWithFormula>().ToArray();
        Assert.Equal(arrayWithFormulas.Length, readArray.Length);

        foreach (var item in readArray)
        {
            Assert.NotNull(item.FullName);
            Assert.Equal($"{item.Name}{item.Surname}", item.FullName.Value);
        }
        
        File.Delete(testFile);
    }

    private static readonly TestNumbersWithFormula[] arrayNumbersWithFormulas =
    {
        new(){ First = 1, Second = 2, Sum = new Formula{FormulaA1 = "SUM(A2:B2)"} },
        new(){ First = 2, Second = 3, Sum = new Formula{FormulaA1 = "SUM(A3:B3)"} },
    };

    [Fact]
    public void NumbersWithFormula()
    {
        var testFile = Path.GetRandomFileName();
        testFile = Path.ChangeExtension(testFile, "xlsx");

        var writer = new ExcelWriter(testFile);
        writer.Write(arrayNumbersWithFormulas);
        writer.SaveAs(testFile);

        var reader = new ExcelReader(testFile);
        var readArray = reader.Read<TestNumbersWithFormula>().ToArray();
        Assert.Equal(arrayNumbersWithFormulas.Length, readArray.Length);

        foreach (var item in readArray)
        {
            Assert.NotNull(item.Sum);
            Assert.Equal(item.First + item.Second, item.Sum.Value);
        }

        File.Delete(testFile);
    }

    private static readonly TestWithHyperlink[] arrayWithHyperlinks =
    {
        new() { Name = "Bilbo", Surname = "Baggins", Job = "Eater", Link = new Hyperlink{ Value = "Wiki", Link = new XLHyperlink("https://en.wikipedia.org/wiki/Bilbo_Baggins") } },
        new() { Name = "John", Job = "Policeman", Link = new Hyperlink{ Value = "CNN", Link = new XLHyperlink("https://edition.cnn.com/2023/12/10/us/john-okeefe-boston-police-death-cec/index.html") } },
        new() { Name = "Bruce", Surname = "Lee", Job = "Fighter", Link = new Hyperlink{ Value = "IMDb", Link = new XLHyperlink("https://www.imdb.com/name/nm0000045/") } }
    };

    [Fact]
    public void WriteWithHyperlink()
    {
        var testFile = Path.GetRandomFileName();
        testFile = Path.ChangeExtension(testFile, "xlsx");

        var writer = new ExcelWriter(testFile);
        writer.Write(arrayWithHyperlinks);
        writer.SaveAs(testFile);

        var reader = new ExcelReader(testFile);
        var readArray = reader.Read<TestWithHyperlink>().ToArray();
        Assert.Equal(arrayWithFormulas.Length, readArray.Length);

        for (var i = 0; i < readArray.Length; i++)
        {
            Assert.NotNull(arrayWithHyperlinks[i].Link);
            Assert.NotNull(readArray[i].Link);
            Assert.Equal(arrayWithHyperlinks[i].Link?.Value, readArray[i].Link?.Value);
            Assert.Equal(arrayWithHyperlinks[i].Link?.Link?.ToString(), readArray[i].Link?.Link?.ToString());
        }
        
        File.Delete(testFile);
    }
}