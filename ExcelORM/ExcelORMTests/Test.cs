using ExcelORM;

namespace ExcelORMTests;

public record Test
{
    [Column("First name" )]
    public string? Name { get; set; }

    [Column("Last name")]
    public string? Surname { get; set; }

    [Column(new[]{"Occupation", "Job"})]
    public string? Job { get; set; }
}