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
        var type = typeof(T);
        foreach (var row in rows)
        {
            if (SkipHidden && row.IsHidden) continue;

            var current = new T();
            foreach (var item in mapping)
            {
                if (item.Position == null || item.PropertyName == null) continue;

                var cell = row.Cell(item.Position.Value);
                if (cell == null || cell.Value.IsBlank) continue;

                var property = type.GetProperty(item.PropertyName);
                if (property == null) continue;

                current.SetPropertyValue(property, cell.Value);
            }

            yield return current;
        }
    }

    public IEnumerable<T> Read<T>(uint startFrom = 1, uint skip = 0) where T : class, new()
    {
        return xlWorkbook.Worksheets.SelectMany(worksheet => Read<T>(worksheet, startFrom, skip));
    }

    public IEnumerable<T> Read<T>(string? worksheetName, uint startFrom = 1, uint skip = 0) where T : class, new()
    {
        var worksheet = xlWorkbook.Worksheets.FirstOrDefault(x => x.Name.Equals(worksheetName, StringComparison.InvariantCultureIgnoreCase));
        if (worksheet == null) yield break;

        foreach (var value in Read<T>(worksheet, startFrom, skip))
            yield return value;
    }

    private IEnumerable<T> Read<T>(IXLWorksheet? worksheet, uint startFrom, uint skip) where T : class, new()
    {
        if (worksheet == null) yield break;

        var firstRow = worksheet.Row((int)startFrom);
        if (firstRow.IsEmpty())
            firstRow = worksheet.RowsUsed().First(x => x.RowNumber() > startFrom && !x.IsEmpty());
        
        var mapping = Mapping.MapProperties<T>(firstRow.CellsUsed());
        if (mapping == null) yield break;

        var rowsToProcess = (ObeyFilter && worksheet.AutoFilter.IsEnabled) switch
        {
            true => worksheet.AutoFilter.VisibleRows
                .Where(x => x.RowNumber() > firstRow.RowNumber())
                .Select(x => x.WorksheetRow()),
            false => worksheet.RowsUsed().Where(x => x.RowNumber() > firstRow.RowNumber())
                
        };

        rowsToProcess = rowsToProcess
            .Skip((int)skip);
        
        foreach (var item in ProcessRows<T>(rowsToProcess, mapping))
            yield return item;
    } 
}