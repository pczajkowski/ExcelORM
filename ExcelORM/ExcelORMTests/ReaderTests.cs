namespace ExcelORMTests;

public class ReaderTests
{
    private const string RegularFile = "testFiles/first.xlsx";
    private const string RegularFileNoAttributes = "testFiles/firstNoAttributes.xlsx";
    private const string HiddenFile = "testFiles/hidden.xlsx";
    private const string FilteredFile = "testFiles/filtered.xlsx";
    private const string DifficultFile = "testFiles/columnsOnTheLeftHeaderNotFirstRow.xlsx";
    private const string MultipleSheetsFile = "testFiles/multipleSheets.xlsx";
    private const string DifferentTypesFile = "testFiles/differentTypes.xlsx";
    private const string WithFormulaFile = "testFiles/withFormula.xlsx";
    private const string AdditionalTypes = "testFiles/additionalTypes.xlsx";
    private const string FallBackToPropertyName = "testFiles/fallBackToPropertyName.xlsx";

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
    public void ReadAdditionalTypes()
    {
        using var reader = new ExcelReader(AdditionalTypes);
        var results = reader.Read<TestAdditionalTypes>().ToArray();
        Assert.NotEmpty(results);
    }

    [Fact]
    public void ReadAdditionalTypesNullable()
    {
        using var reader = new ExcelReader(AdditionalTypes);
        var results = reader.Read<TestAdditionalTypesNullable>().ToArray();
        Assert.NotEmpty(results);
    }

    private const string EmptyFileForAppend = "testFiles/emptyForAppend.xlsx";

    [Fact]
    public void ReadEmptyFile()
    {
        using var reader = new ExcelReader(EmptyFileForAppend);
        var results = reader.Read<Test>().ToArray();
        Assert.Empty(results);
    }

    [Fact]
    public void ReadData()
    {
        var results = ExcelReader.ReadData<Test>(RegularFile).ToArray();
        Assert.NotEmpty(results);

        using var workbook = new XLWorkbook(RegularFile);
        var resultsWorkbook = ExcelReader.ReadData<Test>(workbook).ToArray();
        Assert.NotEmpty(resultsWorkbook);
        Assert.Equal(results, resultsWorkbook);
    }

    [Fact]
    public void FallbackToPropertyName()
    {
        using var reader = new ExcelReader(FallBackToPropertyName);
        var resultsNoFallback = reader.Read<Test>().ToArray();
        Assert.NotEmpty(resultsNoFallback);

        reader.FallbackToPropertyName = true;
        var resultsFallback = reader.Read<Test>().ToArray();
        Assert.NotEmpty(resultsFallback);

        for (var i = 0; i < resultsNoFallback.Length; i++)
        {
            Assert.NotEqual(resultsNoFallback[i].Name, resultsFallback[i].Name);
            Assert.Equal(resultsNoFallback[i].Surname, resultsFallback[i].Surname);
            Assert.Equal(resultsNoFallback[i].Job, resultsFallback[i].Job);
        }
    }
    
    [Fact]
    public void ReadNoAttributes()
    {
        using var reader = new ExcelReader(RegularFile);
        var results = reader.Read<Test>().ToArray();
        Assert.NotEmpty(results);
        
        using var readerNoAttributes = new ExcelReader(RegularFileNoAttributes);
        var resultsNoAttributes = readerNoAttributes.Read<TestNoAttributes>().ToArray();
        Assert.NotEmpty(resultsNoAttributes);
        
        for (var i = 0; i < results.Length; i++)
        {
            Assert.Equal(results[i].Name, resultsNoAttributes[i].name);
            Assert.Equal(results[i].Surname, resultsNoAttributes[i].SurName);
            Assert.Equal(results[i].Job, resultsNoAttributes[i].joB);
        }
    }
    
    [Fact]
    public void ReadNoAttributesWithSkip()
    {
        using var reader = new ExcelReader(RegularFile);
        var results = reader.Read<TestWithSkip>().ToArray();
        Assert.NotEmpty(results);
        
        using var readerNoAttributes = new ExcelReader(RegularFileNoAttributes);
        var resultsNoAttributes = readerNoAttributes.Read<TestNoAttributesWithSkip>().ToArray();
        Assert.NotEmpty(resultsNoAttributes);
        
        for (var i = 0; i < results.Length; i++)
        {
            Assert.Equal(results[i].Name, resultsNoAttributes[i].name);
            
            Assert.Equal(results[i].Surname, resultsNoAttributes[i].SurName);
            Assert.Null(results[i].Surname);
            Assert.Null(resultsNoAttributes[i].SurName);
            
            Assert.Equal(results[i].Job, resultsNoAttributes[i].joB);
        }
    }
    
    [Fact]
    public void ReadIncludeQuotePrefix()
    {
        using var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add("Sheet1");
        
        worksheet.Cell(1, 1).Value = nameof(Test.Name);
        worksheet.Cell(2, 1).Value = "Test";
        worksheet.Cell(2, 1).Style.SetIncludeQuotePrefix();

        var reader = new ExcelReader(workbook);
        reader.FallbackToPropertyName = true;
        var results = reader.Read<Test>().ToArray();
        Assert.Equal("'Test", results[0].Name);
    }
}