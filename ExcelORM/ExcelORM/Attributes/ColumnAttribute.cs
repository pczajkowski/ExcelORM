namespace ExcelORM.Attributes
{
    [AttributeUsage(AttributeTargets.Property)]
    public class ColumnAttribute : Attribute
    {
        public string[]? Names { get; init; }

        public ColumnAttribute(string name)
        {
            Names = new[] { name };
        }

        public ColumnAttribute(string[] names)
        {
            Names = names;
        }
    }
}
