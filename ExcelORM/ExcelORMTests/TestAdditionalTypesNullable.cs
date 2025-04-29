namespace ExcelORMTests;

public record TestAdditionalTypesNullable
{
    public TestEnum? MyEnum {get; set;}
    public Guid? MyGuid {get; set;}
}