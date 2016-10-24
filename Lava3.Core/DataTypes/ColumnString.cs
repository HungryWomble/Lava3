
using System.Collections.Generic;
using OfficeOpenXml;

namespace Lava3.Core.DataTypes
{
    public class ColumnString
    {
        public ColumnString()
        {
            Errors = new List<string>();
        }

        public ColumnString(ExcelWorksheet sheet, int rownum, dynamic categoryColumns)
        {
            Errors = new List<string>();
            ColumnNumber = categoryColumns.ColumnNumber;
            Value = Common.ReplaceNullOrEmpty(sheet.Cells[rownum, ColumnNumber].Value).ToString(); ;
        }

        public string Value { get; set; }
        public int ColumnNumber { get; set; }

        public override string ToString()
        {
            return $"[{ColumnNumber}] {Value}";
        }
        public List<string> Errors { get; set; }
    }
}
