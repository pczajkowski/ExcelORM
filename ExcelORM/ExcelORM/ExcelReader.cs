using ClosedXML.Excel;

namespace ExcelORM;

public class ExcelReader
{
    private readonly IXLWorkbook xlWorkbook;
    public uint SkipFirstNRows { get; set; } = 1;
    public bool SkipHidden { get; set; }
    public bool ObeyFilter { get; set; }

    public ExcelReader(string? path)
    {
        xlWorkbook = new XLWorkbook(path);
    }

    private IEnumerable<T> ProcessRows<T>(IEnumerable<IXLRow> rows, List<Mapping> mapping) where T : class, new()
    {
        foreach (var row in rows)
        {
            if (SkipHidden && row.IsHidden) continue;

            var current = new T();
            foreach (var item in mapping)
            {
                if (item.Position == null || item.PropertyName == null) continue;

                var cell = row.Cell(item.Position.Value);
                if (cell == null || cell.Value.IsBlank) continue;

                var property = current.GetType().GetProperty(item.PropertyName);
                if (property == null) continue;

                object? valueToSet = property.PropertyType switch
                {
                    not null when property.PropertyType == typeof(string) => cell.Value.ToString(),
                    not null when property.PropertyType == typeof(DateTime?) => cell.Value.IsDateTime ? cell.Value.GetDateTime() : null,
                    not null when property.PropertyType == typeof(TimeSpan?) => cell.Value.IsTimeSpan ? cell.Value.GetTimeSpan() : null,
                    not null when property.PropertyType == typeof(double?) => cell.Value.IsNumber ? cell.Value.GetNumber() : null,
                    not null when property.PropertyType == typeof(int?) => cell.Value.IsNumber ? (int?)cell.Value.GetNumber() : null,
                    _ => throw new NotSupportedException($"{property.PropertyType} isn't supported!")
                };
               
                if (valueToSet != null)
                    property.SetValue(current, valueToSet);
            }

            yield return current;
        }
    }

    public IEnumerable<T> Read<T>() where T : class, new()
    {
        return xlWorkbook.Worksheets.SelectMany(Read<T>);
    }

    public IEnumerable<T> Read<T>(string? worksheetName) where T : class, new()
    {
        var worksheet = xlWorkbook.Worksheets.FirstOrDefault(x => x.Name.Equals(worksheetName, StringComparison.InvariantCultureIgnoreCase));
        if (worksheet == null) yield break;

        foreach (var value in Read<T>(worksheet))
            yield return value;
    }

    private IEnumerable<T> Read<T>(IXLWorksheet? worksheet) where T : class, new()
    {
        if (worksheet == null) yield break;

        var mapping = Mapping.MapProperties<T>(worksheet.FirstRowUsed().CellsUsed());
        if (mapping == null) yield break;

        if (ObeyFilter && worksheet.AutoFilter.IsEnabled)
        {
            foreach (var item in ProcessRows<T>(worksheet.AutoFilter.VisibleRows
                         .Select(x => x.WorksheetRow()).Skip((int)SkipFirstNRows), mapping))
                yield return item;
        }
        else
        {
            foreach (var item in ProcessRows<T>(worksheet.RowsUsed().Skip((int)SkipFirstNRows), mapping))
                yield return item;
        }
    } 
}