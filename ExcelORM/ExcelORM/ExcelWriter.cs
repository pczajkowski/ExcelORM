using ClosedXML.Excel;
using ExcelORM.Attributes;
using ExcelORM.Interfaces;
using ExcelORM.Models;

namespace ExcelORM;

public class ExcelWriter
{
    private readonly IXLWorkbook xlWorkbook;
    public ExcelWriter(string? path = null)
    {
        xlWorkbook = File.Exists(path) ? new XLWorkbook(path) : new XLWorkbook();
    }

    private static int GenerateHeader<T>(IXLWorksheet worksheet) where T : class
    {
        var rowIndex = 1;
        var cellIndex = 1;
        var properties = typeof(T).GetProperties();
        foreach (var property in properties)
        {
            if (property.Skip()) continue;

            var columnAttribute = property.GetCustomAttributes(typeof(ColumnAttribute), false).FirstOrDefault() as ColumnAttribute;
            worksheet.Cell(rowIndex, cellIndex).Value = columnAttribute is { Names.Length: > 0 } ? columnAttribute.Names.First() : property.Name;
            cellIndex++;
        }

        return ++rowIndex;
    }

    private static void Write<T>(IEnumerable<T> values, IXLWorksheet worksheet, bool append) where T : class
    {
        if (!values.Any()) return;

        var rowIndex = append switch
        {
            true => worksheet.LastRowUsed().RowNumber() + 1,
            false => GenerateHeader<T>(worksheet),
        };

        foreach (var value in values)
        {
            var cellIndex = 0;
            var properties = typeof(T).GetProperties();
            foreach (var property in properties)
            {
                if (property.Skip()) continue;

                cellIndex++;

                var valueToSet = property.GetValue(value);
                if (valueToSet == null) continue;
                
                if (valueToSet is ISpecialProperty specialProperty)
                {
                    specialProperty.SetCellValue(worksheet.Cell(rowIndex, cellIndex));
                    continue;
                }

                worksheet.Cell(rowIndex, cellIndex).Value = XLCellValue.FromObject(valueToSet);
            }

            rowIndex++;
        }
    }

    public void Write<T>(IEnumerable<T> values, string? worksheetName = null, bool append = false) where T : class
    {
        var xlWorksheet = xlWorkbook.Worksheets.FirstOrDefault(x => x.Name.Equals(worksheetName, StringComparison.InvariantCultureIgnoreCase));
        
        xlWorksheet ??= !string.IsNullOrWhiteSpace(worksheetName) ?
            xlWorkbook.AddWorksheet(worksheetName)
            : xlWorkbook.Worksheets.Count == 0 ? xlWorkbook.AddWorksheet() : xlWorkbook.Worksheets.First();

        Write(values, xlWorksheet, append);
    }

    public void SaveAs(string path, IExcelConverter? converter = null)
    {
        xlWorkbook.SaveAs(path);
        converter?.MakeCompatible(path);
    } 
}