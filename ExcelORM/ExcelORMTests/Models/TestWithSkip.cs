using ExcelORM.Attributes;

namespace ExcelORMTests;

public record TestWithSkip
{
    [Column("first name" )]
    public string? Name { get; set; }

    [Skip]
    [Column("Last Name")]
    public string? Surname { get; set; }

    [Column(["Occupation", "JoB"])]
    public string? Job { get; set; }
}