using ClosedXML.Excel;

namespace ExcelORM;

public class ExcelReader
{
    private readonly IXLWorkbook xlWorkbook;
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

                current.SetValue(property, cell);
            }

            yield return current;
        }
    }

    public IEnumerable<T> Read<T>(uint startFrom = 1, uint skip = 1) where T : class, new()
    {
        return xlWorkbook.Worksheets.SelectMany(worksheet => Read<T>(worksheet, startFrom, skip));
    }

    public IEnumerable<T> Read<T>(string? worksheetName, uint startFrom = 1, uint skip = 1) where T : class, new()
    {
        var worksheet = xlWorkbook.Worksheets.FirstOrDefault(x => x.Name.Equals(worksheetName, StringComparison.InvariantCultureIgnoreCase));
        if (worksheet == null) yield break;

        foreach (var value in Read<T>(worksheet, startFrom, skip))
            yield return value;
    }

    private IEnumerable<T> Read<T>(IXLWorksheet? worksheet, uint startFrom, uint skip) where T : class, new()
    {
        if (worksheet == null) yield break;

        var mapping = Mapping.MapProperties<T>(worksheet.FirstRowUsed().CellsUsed());
        if (mapping == null) yield break;

        var rowsToProcess = (ObeyFilter && worksheet.AutoFilter.IsEnabled) switch
        {
            true => worksheet.AutoFilter.VisibleRows
                .Select(x => x.WorksheetRow()),
            false => worksheet.RowsUsed()
                
        };

        rowsToProcess = rowsToProcess.Where(x => x.RowNumber() >= startFrom)
            .Skip((int)skip);
        
        foreach (var item in ProcessRows<T>(rowsToProcess, mapping))
            yield return item;
    } 
}