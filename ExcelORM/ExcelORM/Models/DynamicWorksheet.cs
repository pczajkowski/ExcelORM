namespace ExcelORM.Models
{
    public record DynamicWorksheet
    {
        public string? Name { get; set; }
        public int Position { get; set; }
        public IEnumerable<List<DynamicCell>>? Cells { get; set; }
    }
}
