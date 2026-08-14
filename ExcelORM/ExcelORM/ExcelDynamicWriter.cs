using ClosedXML.Excel;
using ExcelORM.Interfaces;
using ExcelORM.Models;

namespace ExcelORM;

public sealed class ExcelDynamicWriter : IDisposable
{
    private readonly IXLWorkbook xlWorkbook;
    public ExcelDynamicWriter(string? path = null)
    {
        xlWorkbook = File.Exists(path) ? new XLWorkbook(path) : new XLWorkbook();
    }

    public ExcelDynamicWriter(IXLWorkbook? workbook)
    {
        xlWorkbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
    }

    private static int GenerateHeader(IXLWorksheet worksheet, IEnumerable<DynamicCell> firstRow)
    {
        var rowIndex = 1;
        foreach (var item in firstRow)
            worksheet.Cell(rowIndex, item.Position).Value = item.Header;

        return ++rowIndex;
    }

    private static void Write(IEnumerable<List<DynamicCell>> values, IXLWorksheet worksheet, bool append)
    {
        var valuesList = values.ToList();
        if (valuesList.Count == 0) return;
        
        var lastRow = worksheet.LastRowUsed();
        if (lastRow == null)
        {
            append = false;
            lastRow = worksheet.FirstRow();
        }

        var rowIndex = append switch
        {
            true => lastRow.RowNumber() + 1,
            false => GenerateHeader(worksheet, valuesList.First()),
        };

        foreach (var row in valuesList)
        {
            foreach (var cell in row)
            {
                if (cell.Value == null) continue;

                worksheet.Cell(rowIndex, cell.Position).Value = XLCellValue.FromObject(cell.Value);
            }

            rowIndex++;
        }
    }

    public void Write(IEnumerable<List<DynamicCell>>? values, string? worksheetName = null, bool append = false)
    {
        if (values == null) return;

        var xlWorksheet = xlWorkbook.Worksheets.FirstOrDefault(x => x.Name.Equals(worksheetName, StringComparison.InvariantCultureIgnoreCase));
        
        xlWorksheet ??= !string.IsNullOrWhiteSpace(worksheetName) ?
            xlWorkbook.AddWorksheet(worksheetName)
            : xlWorkbook.Worksheets.Count == 0 ? xlWorkbook.AddWorksheet() : xlWorkbook.Worksheets.First();

        Write(values, xlWorksheet, append);
    }

    public void WriteAll(IEnumerable<DynamicWorksheet>? dynamicWorksheets, bool append = false)
    {
        if (dynamicWorksheets == null) return;

        foreach (var dynamicWorksheet in dynamicWorksheets)
            Write(dynamicWorksheet.Cells, dynamicWorksheet.Name, append);
    }

    public void SaveAs(string path, IExcelConverter? converter = null)
    {
        xlWorkbook.SaveAs(path);
        converter?.MakeCompatible(path);
    }

    public void Dispose()
    {
        xlWorkbook.Dispose();
    }
}