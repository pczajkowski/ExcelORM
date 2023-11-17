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
            if (row.RowNumber() <= SkipFirstNRows) continue;

            var current = new T();
            foreach (var item in mapping)
            {
                if (item.Position == null || item.PropertyName == null) continue;

                var cell = row.Cell(item.Position.Value);
                if (cell == null || cell.Value.IsBlank) continue;

                var property = current.GetType().GetProperty(item.PropertyName);
                if (property == null) continue;

                switch (property.PropertyType)
                {
                    case not null when property.PropertyType == typeof(string):
                        property.SetValue(current, cell.Value.ToString());
                        break;
                }
            }

            yield return current;
        }
    }

    public IEnumerable<T> Read<T>() where T : class, new()
    {
        foreach (var worksheet in xlWorkbook.Worksheets)
        {
            foreach (var value in Read<T>(worksheet))
                yield return value;
        }
    }

    public IEnumerable<T> Read<T>(string? worksheetName) where T : class, new()
    {
        var worksheet = xlWorkbook.Worksheets.FirstOrDefault(x => x.Name.Equals(worksheetName, StringComparison.InvariantCultureIgnoreCase));
        if (worksheet == null) yield break;

        foreach (var value in Read<T>(worksheet))
            yield return value;
    }

    public IEnumerable<T> Read<T>(IXLWorksheet? worksheet) where T : class, new()
    {
        if (worksheet == null) yield break;

        var mapping = Mapping.MapProperties<T>(worksheet.FirstRowUsed().CellsUsed());
        if (mapping == null) yield break;

        if (ObeyFilter && worksheet.AutoFilter.IsEnabled)
        {
            foreach (var item in ProcessRows<T>(worksheet.AutoFilter.VisibleRows.Select(x => x.WorksheetRow()), mapping))
                yield return item;
        }
        else
        {
            foreach (var item in ProcessRows<T>(worksheet.RowsUsed(), mapping))
                yield return item;
        }
    } 
}