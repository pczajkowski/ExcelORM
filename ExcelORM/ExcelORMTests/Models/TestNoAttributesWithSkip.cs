using ExcelORM.Attributes;

namespace ExcelORMTests;

public record TestNoAttributesWithSkip
{
    public string? name { get; set; }

    [Skip]
    public string? SurName { get; set; }

    public string? joB { get; set; }
}