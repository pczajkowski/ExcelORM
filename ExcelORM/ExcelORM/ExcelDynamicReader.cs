using ClosedXML.Excel;
using ExcelORM.Models;

namespace ExcelORM;

public class ExcelDynamicReader : IDisposable
{
    private readonly IXLWorkbook xlWorkbook;
    public bool SkipHidden { get; set; }
    public bool ObeyFilter { get; set; }

    public ExcelDynamicReader(string? path)
    {
        xlWorkbook = new XLWorkbook(path);
    }

    public ExcelDynamicReader(IXLWorkbook workbook)
    {
        xlWorkbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
    }

    private IEnumerable<List<DynamicCell>> ProcessRows(IEnumerable<IXLRow> rows, List<DynamicCell> mapping)
    {
        foreach (var row in rows)
        {
            if (SkipHidden && row.IsHidden) continue;

            var dynamicRow = new List<DynamicCell>();
            foreach (var item in mapping)
            {
                var cell = row.Cell(item.Position);
                if (cell == null || cell.Value.IsBlank)
                {
                    dynamicRow.Add(item);
                    continue;
                }

                if (item.Type == null) item.Type = cell.Value.ValueType();

                var cellItem = item with
                {
                    Value = cell.Value.ToObject()
                };

                dynamicRow.Add(cellItem);
            }

            yield return dynamicRow;
        }
    }

    private IEnumerable<List<DynamicCell>> Read(IXLWorksheet? worksheet, uint startFrom = 1, uint skip = 0)
    {
        if (worksheet == null) yield break;

        var firstRow = worksheet.Row((int)startFrom);
        if (firstRow.IsEmpty())
            firstRow = worksheet.RowsUsed().First(x => x.RowNumber() > startFrom && !x.IsEmpty());

        var mapping = DynamicCell.MapHeader(firstRow.CellsUsed());
        if (mapping == null || mapping.Count == 0) yield break;

        var rowsToProcess = (ObeyFilter && worksheet.AutoFilter.IsEnabled) switch
        {
            true => worksheet.AutoFilter.VisibleRows
                .Where(x => x.RowNumber() > firstRow.RowNumber())
                .Select(x => x.WorksheetRow()),
            false => worksheet.RowsUsed().Where(x => x.RowNumber() > firstRow.RowNumber())

        };

        rowsToProcess = rowsToProcess
            .Skip((int)skip);

        foreach (var item in ProcessRows(rowsToProcess, mapping))
            yield return item;
    }

    public IEnumerable<List<DynamicCell>> Read(string? worksheetName, uint startFrom = 1, uint skip = 0)
    {
        var worksheet = xlWorkbook.Worksheets.FirstOrDefault(x => x.Name.Equals(worksheetName, StringComparison.InvariantCultureIgnoreCase));
        if (worksheet == null) yield break;

        foreach (var value in Read(worksheet, startFrom, skip))
            yield return value;
    }

    public IEnumerable<List<DynamicCell>> Read(int worksheetIndex = 1, uint startFrom = 1, uint skip = 0)
    {
        if (worksheetIndex > xlWorkbook.Worksheets.Count) yield break;

        var worksheet = xlWorkbook.Worksheets.FirstOrDefault(x => x.Position == worksheetIndex);
        if (worksheet == null) yield break;

        foreach (var value in Read(worksheet, startFrom, skip))
            yield return value;
    }

    public IEnumerable<DynamicWorksheet> ReadAll(uint startFrom = 1, uint skip = 0)
    {
        foreach (var worksheet in xlWorkbook.Worksheets)
        {
            yield return new DynamicWorksheet
            {
                Name = worksheet.Name,
                Position = worksheet.Position,
                Cells = Read(worksheet, startFrom, skip)
            };
        }
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
            xlWorkbook?.Dispose();
        }
    }
    ~ExcelDynamicReader() => Dispose(false);
}