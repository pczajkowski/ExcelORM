using ClosedXML.Excel;

namespace ExcelORM;

public class ExcelWriter
{
    private readonly IXLWorkbook xlWorkbook;
    public bool WriteHeader { get; set; } = true;

    public ExcelWriter(string? path = null)
    {
        xlWorkbook = File.Exists(path) ? new XLWorkbook(path) : new XLWorkbook();
    }

    private static uint GenerateHeader<T>(T value, IXLWorksheet worksheet, uint rowIndex = 1) where T : class, new()
    {
        var cellIndex = 1;
        var properties = value.GetType().GetProperties();
        foreach (var property in properties)
        {
            var columnAttribute = property.GetCustomAttributes(typeof(ColumnAttribute), false).FirstOrDefault() as ColumnAttribute;
            worksheet.Cell((int)rowIndex, cellIndex).Value = columnAttribute is { Names.Length: > 0 } ? columnAttribute.Names.First() : property.Name;
            cellIndex++;
        }

        return ++rowIndex;
    }

    public void Write<T>(IEnumerable<T> values, string? worksheetName, bool append = false, uint rowIndex = 1) where T : class, new()
    {
        var xlWorksheet = xlWorkbook.Worksheets.FirstOrDefault(x => x.Name.Equals(worksheetName, StringComparison.InvariantCultureIgnoreCase));
        xlWorksheet ??= !string.IsNullOrWhiteSpace(worksheetName) ?
            xlWorkbook.AddWorksheet(worksheetName) : xlWorkbook.AddWorksheet();

        Write(values, xlWorksheet, append, rowIndex);
    }

    private void Write<T>(IEnumerable<T> values, IXLWorksheet worksheet, bool append = false, uint rowIndex = 1) where T : class, new()
    {
        var enumerable = values as T[] ?? values.ToArray();
        if (!enumerable.Any()) return;

        rowIndex = append switch
        {
            true => (uint)worksheet.LastRowUsed().RowNumber() + 1,
            false when WriteHeader => GenerateHeader(enumerable.First(), worksheet),
            _ => rowIndex
        };

        foreach (var value in enumerable)
        {
            var cellIndex = 1;
            var properties = value.GetType().GetProperties();
            foreach (var property in properties)
            {
                worksheet.Cell((int)rowIndex, cellIndex).Value = property.GetValue(value) as string;
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