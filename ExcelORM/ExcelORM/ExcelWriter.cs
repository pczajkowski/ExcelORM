using ClosedXML.Excel;

namespace ExcelORM;

public class ExcelWriter
{
    private readonly IXLWorkbook xlWorkbook;
    public bool WriteHeader { get; set; } = true;

    public ExcelWriter(string? path = null)
    {
        if (File.Exists(path))
            xlWorkbook = new XLWorkbook(path);
        else
            xlWorkbook = new XLWorkbook();
    }

    private uint GenerateHeader<T>(T value, IXLWorksheet worksheet, uint rowIndex = 1) where T : class, new()
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

    public void Write<T>(IEnumerable<T> values, IXLWorksheet worksheet, bool append = false, uint rowIndex = 1) where T : class, new()
    {
        if (!values.Any()) return;

        if (append)
            rowIndex = (uint)worksheet.LastRowUsed().RowNumber() + 1;

        if (!append && WriteHeader)
            rowIndex = GenerateHeader(values.First(), worksheet);

        foreach (var value in values)
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