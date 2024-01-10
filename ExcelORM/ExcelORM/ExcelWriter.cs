using ClosedXML.Excel;
using ExcelORM.Attributes;

namespace ExcelORM;

public class ExcelWriter
{
    private readonly IXLWorkbook xlWorkbook;
    public ExcelWriter(string? path = null)
    {
        xlWorkbook = File.Exists(path) ? new XLWorkbook(path) : new XLWorkbook();
    }

    private static int GenerateHeader<T>(IXLWorksheet worksheet) where T : class, new()
    {
        var rowIndex = 1;
        var cellIndex = 1;
        var properties = typeof(T).GetProperties();
        foreach (var property in properties)
        {
            var columnAttribute = property.GetCustomAttributes(typeof(ColumnAttribute), false).FirstOrDefault() as ColumnAttribute;
            worksheet.Cell(rowIndex, cellIndex).Value = columnAttribute is { Names.Length: > 0 } ? columnAttribute.Names.First() : property.Name;
            cellIndex++;
        }

        return ++rowIndex;
    }

    public void Write<T>(IEnumerable<T> values, string? worksheetName = null, bool append = false) where T : class, new()
    {
        var xlWorksheet = xlWorkbook.Worksheets.FirstOrDefault(x => x.Name.Equals(worksheetName, StringComparison.InvariantCultureIgnoreCase));
        
        xlWorksheet ??= !string.IsNullOrWhiteSpace(worksheetName) ?
            xlWorkbook.AddWorksheet(worksheetName)
            : xlWorkbook.Worksheets.Count == 0 ? xlWorkbook.AddWorksheet() : xlWorkbook.Worksheets.First();

        Write(values, xlWorksheet, append);
    }

    private static void Write<T>(IEnumerable<T> values, IXLWorksheet worksheet, bool append) where T : class, new()
    {
        var enumerable = values as T[] ?? values.ToArray();
        if (!enumerable.Any()) return;

        var rowIndex = append switch
        {
            true => worksheet.LastRowUsed().RowNumber() + 1,
            false => GenerateHeader<T>(worksheet),
        };

        foreach (var value in enumerable)
        {
            var cellIndex = 1;
            var properties = typeof(T).GetProperties();
            foreach (var property in properties)
            {
                var valueToSet = property.GetValue(value);
                if (valueToSet == null) continue;
                
                worksheet.Cell(rowIndex, cellIndex).Value = XLCellValue.FromObject(valueToSet);
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