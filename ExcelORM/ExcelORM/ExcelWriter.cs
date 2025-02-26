using System.Reflection;
using ClosedXML.Excel;
using ExcelORM.Attributes;
using ExcelORM.Interfaces;
using ExcelORM.Models;

namespace ExcelORM;

public class ExcelWriter : IDisposable
{
    private readonly IXLWorkbook xlWorkbook;
    public ExcelWriter(string? path = null)
    {
        xlWorkbook = File.Exists(path) ? new XLWorkbook(path) : new XLWorkbook();
    }

    public ExcelWriter(IXLWorkbook workbook)
    {
        xlWorkbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
    }

    private static int GenerateHeader(IXLWorksheet worksheet, PropertyInfo[] properties)
    {
        var rowIndex = 1;
        var cellIndex = 1;
        foreach (var property in properties)
        {
            if (property.Skip()) continue;

            var columnAttribute = property.GetCustomAttributes(typeof(ColumnAttribute), false).FirstOrDefault() as ColumnAttribute;
            worksheet.Cell(rowIndex, cellIndex).Value = columnAttribute is { Names.Length: > 0 } ? columnAttribute.Names.First() : property.Name;
            cellIndex++;
        }

        return ++rowIndex;
    }

    private static void WriteCell<T>(T value, PropertyInfo property, IXLCell cell)
    {
        var valueToSet = property.GetValue(value);
        if (valueToSet == null) return;

        if (valueToSet is SpecialBase specialProperty)
        {
            specialProperty.SetCellValue(cell);
            return;
        }

        cell.Value = XLCellValue.FromObject(valueToSet);
    }

    private static void WriteRowAppend<T>(T value, IXLWorksheet worksheet, PropertyInfo[] properties, int rowIndex, List<Mapping> mapping)
    {
        foreach (var property in properties)
        {
            if (property.Skip()) continue;

            var mapped = mapping.FirstOrDefault(x => x.PropertyName != null && x.PropertyName.Equals(property.Name));
            if (mapped?.Position == null) continue;

            WriteCell(value, property, worksheet.Cell(rowIndex, mapped.Position.Value));
        }
    }

    private static void WriteRow<T>(T value, IXLWorksheet worksheet, PropertyInfo[] properties, int rowIndex)
    {
        var cellIndex = 0;
        foreach (var property in properties)
        {
            if (property.Skip()) continue;

            cellIndex++;

            WriteCell(value, property, worksheet.Cell(rowIndex, cellIndex));
        }
    }

    private static void Write<T>(IEnumerable<T> values, IXLWorksheet worksheet, bool append, uint? headerRowIndex = null, uint? appendFrom = null) where T : class
    {
        var properties = typeof(T).GetProperties();
        List<Mapping>? mapping = [];

        var rowIndex = (append, startFrom: appendFrom) switch
        { 
            (true, not null) => (int)appendFrom,
            (true, null) => worksheet.LastRowUsed()?.RowNumber() + 1,
            _ => GenerateHeader(worksheet, properties) 
        };

        if (append)
        {
            var headerCells = headerRowIndex != null ? worksheet.Row((int)headerRowIndex).CellsUsed() : worksheet.FirstRowUsed()?.CellsUsed();
            mapping = Mapping.MapProperties<T>(headerCells);
            if (mapping == null || mapping.Count == 0) return;
        }

        if (rowIndex == null) throw new NullReferenceException(nameof(rowIndex));

        foreach (var value in values)
        {
            if (append) WriteRowAppend(value, worksheet, properties, rowIndex.Value, mapping);
            else WriteRow(value, worksheet, properties, rowIndex.Value);
            
            rowIndex++;
        }
    }

    public void Write<T>(IEnumerable<T> values, string? worksheetName = null, bool append = false, uint? headerRowIndex = null) where T : class
    {
        var xlWorksheet = xlWorkbook.Worksheets.FirstOrDefault(x => x.Name.Equals(worksheetName, StringComparison.InvariantCultureIgnoreCase));
        
        xlWorksheet ??= !string.IsNullOrWhiteSpace(worksheetName) ?
            xlWorkbook.AddWorksheet(worksheetName)
            : xlWorkbook.Worksheets.Count == 0 ? xlWorkbook.AddWorksheet() : xlWorkbook.Worksheets.First();

        Write(values, xlWorksheet, append, headerRowIndex);
    }

    public void SaveAs(string path, IExcelConverter? converter = null)
    {
        xlWorkbook.SaveAs(path);
        converter?.MakeCompatible(path);
    } 

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    protected virtual void Dispose(bool disposing)
    {
        if (disposing)
        {
            xlWorkbook.Dispose();
        }
    }
    ~ExcelWriter() => Dispose(false);
}