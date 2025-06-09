namespace ExcelORM.Attributes
{
    [AttributeUsage(AttributeTargets.Property)]
    public class SkipAttribute : Attribute
    {
        public bool SkipOnWrite { get; init; }
        public bool SkipOnRead { get; init; }

        public SkipAttribute()
        {
            SkipOnWrite = true;
            SkipOnRead = true;
        }
        public SkipAttribute(bool skipOnWrite, bool skipOnRead)
        {
            SkipOnWrite = skipOnWrite;
            SkipOnRead = skipOnRead;
        }
    }
}
