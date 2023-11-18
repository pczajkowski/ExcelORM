using ClosedXML.Excel;

namespace ExcelORM;

public class ExcelWriter
{
    private readonly IXLWorkbook xlWorkbook;
    public ExcelWriter(string? path = null)
    {
        xlWorkbook = File.Exists(path) ? new XLWorkbook(path) : new XLWorkbook();
    }

    private static int GenerateHeader<T>(T value, IXLWorksheet worksheet) where T : class, new()
    {
        var rowIndex = 1;
        var cellIndex = 1;
        var properties = value.GetType().GetProperties();
        foreach (var property in properties)
        {
            var columnAttribute = property.GetCustomAttributes(typeof(ColumnAttribute), false).FirstOrDefault() as ColumnAttribute;
            worksheet.Cell(rowIndex, cellIndex).Value = columnAttribute is { Names.Length: > 0 } ? columnAttribute.Names.First() : property.Name;
            cellIndex++;
        }

        return ++rowIndex;
    }

    public void Write<T>(IEnumerable<T> values, string? worksheetName, bool append = false) where T : class, new()
    {
        var xlWorksheet = xlWorkbook.Worksheets.FirstOrDefault(x => x.Name.Equals(worksheetName, StringComparison.InvariantCultureIgnoreCase));
        xlWorksheet ??= !string.IsNullOrWhiteSpace(worksheetName) ?
            xlWorkbook.AddWorksheet(worksheetName) : xlWorkbook.AddWorksheet();

        Write(values, xlWorksheet, append);
    }

    private void Write<T>(IEnumerable<T> values, IXLWorksheet worksheet, bool append = false) where T : class, new()
    {
        var enumerable = values as T[] ?? values.ToArray();
        if (!enumerable.Any()) return;

        var rowIndex = append switch
        {
            true => worksheet.LastRowUsed().RowNumber() + 1,
            false => GenerateHeader(enumerable.First(), worksheet),
        };

        foreach (var value in enumerable)
        {
            var cellIndex = 1;
            var properties = value.GetType().GetProperties();
            foreach (var property in properties)
            {
                worksheet.Cell(rowIndex, cellIndex).Value = property.GetValue(value) as string;
                cellIndex++;
            }

            rowIndex++;
        }
    }

    public void SaveAs(string path, IExcelConverter? converter = null)
    {
        xlWorkbook.SaveAs(path);
        converter?.MakeCompatible(path);
    } 
}