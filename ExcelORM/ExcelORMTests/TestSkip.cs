using ExcelORM.Attributes;

namespace ExcelORMTests
{
    public record TestSkip
    {
        [Skip]
        public string? Text { get; set; }
        public DateTime? Date { get; set; }
        public double? Int { get; set; }
    }
}
