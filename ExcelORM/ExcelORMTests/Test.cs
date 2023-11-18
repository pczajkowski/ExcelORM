using ExcelORM;

namespace ExcelORMTests;

public record Test
{
    [Column("first name" )]
    public string? Name { get; set; }

    [Column("Last Name")]
    public string? Surname { get; set; }

    [Column(new[]{"Occupation", "JoB"})]
    public string? Job { get; set; }
}