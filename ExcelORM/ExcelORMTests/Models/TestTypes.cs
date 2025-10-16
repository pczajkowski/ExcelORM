namespace ExcelORMTests;

public record TestTypes
{
    public string? Text { get; set; }
    public DateTime? Date { get; set; }
    public TimeSpan? TimeSpan { get; set; }
    public int? Int { get; set; }
    public double? Double { get; set; }
    public decimal? Decimal { get; set; }
    public long? Long { get; set; }
    public short? Short { get; set; }
    public float? Float { get; set; }
    public byte? Byte { get; set; }
}