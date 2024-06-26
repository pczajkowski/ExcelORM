using ExcelORM.Models;

namespace ExcelORMTests
{
    public record TestNumbersWithFormula
    {
        public double First { get; set; }
        public double Second { get; set; }
        public Formula? Sum { get; set; }
    }
}
