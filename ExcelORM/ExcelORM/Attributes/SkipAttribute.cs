namespace ExcelORM.Attributes
{
    [AttributeUsage(AttributeTargets.Property)]
    public class SkipAttribute : Attribute
    {
        public bool SkipOnWrite { get; init; } = true;
        public bool SkipOnRead { get; init; } = true;
        public SkipAttribute() { }
        public SkipAttribute(bool skipOnWrite, bool skipOnRead)
        {
            SkipOnWrite = skipOnWrite;
            SkipOnRead = skipOnRead;
        }
    }
}
