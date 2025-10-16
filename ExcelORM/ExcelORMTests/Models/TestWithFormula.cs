using ExcelORM.Attributes;
using ExcelORM.Models;

namespace ExcelORMTests
{
    public record TestWithFormula : Test
    {
        [Column("Full name")]
        public Formula? FullName { get; set; }
    }
}
