using ExcelORM.Models;

namespace ExcelORMTests
{
    public record TestWithHyperlink : Test
    {
        public Hyperlink? Link { get; set; }
    }
}
