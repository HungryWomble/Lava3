
namespace Lava3.Core
{
    public class ColumnHeader
    {
        public string Header { get; set; }
        public int ColumnNumber { get; set; }
        public override string ToString()
        {
            return $"[{ColumnNumber}] {Header}";
        }
        public string GetColumnLetter()
        {
            return Common.GetExcelColumnLetter(this.ColumnNumber);
        }
    }
}
