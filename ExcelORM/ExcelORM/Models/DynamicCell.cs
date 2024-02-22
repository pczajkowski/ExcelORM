namespace ExcelORM.Models
{
    public record DynamicCell
    {
        public int Position { get; set; }
        public string? Header { get; set; }
        public Type? Type { get; set; }
        public object? Value { get; set; }
    }
}
