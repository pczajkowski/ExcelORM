using ExcelORM.Attributes;

namespace ExcelORMTests
{
    public record TestSkipMiddle
    {
        public string? Text { get; set; }

        [Skip]
        public DateTime? Date { get; set; }
        public double? Int { get; set; }
    }
}
