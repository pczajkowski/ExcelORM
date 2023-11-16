namespace ExcelORM
{
    [AttributeUsage(AttributeTargets.Property)]
    public class ColumnAttribute : Attribute
    {
        public string[]? Names { get; init; }

        public ColumnAttribute(string name)
        {
            Names = new string[] { name };
        }

        public ColumnAttribute(string[] names)
        {
            Names = names;
        }
    }
}
