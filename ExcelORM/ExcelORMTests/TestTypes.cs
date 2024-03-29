namespace ExcelORMTests;

public record TestTypes
{
    public string? Text { get; set; }
    public DateTime? Date { get; set; }
    public TimeSpan? TimeSpan { get; set; }
    public double? Int { get; set; }
    public double? Double { get; set; }
}