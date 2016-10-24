using System.Collections.Generic;
using OfficeOpenXml;

namespace Lava3.Core.DataTypes
{
    public class ColumnDecimal 
    {
        public ColumnDecimal()
        {
            Errors = new List<string>();
        }

        public ColumnDecimal(ExcelWorksheet sheet, int rownum, dynamic categoryColumns)
        {
            Errors = new List<string>();
            ColumnNumber = categoryColumns.ColumnNumber;
            Value = Common.ReplaceNullOrEmptyDecimal(sheet.Cells[rownum, ColumnNumber].Value);
        }

        public decimal? Value { get; set; }
        public int ColumnNumber { get; set; }

        public override string ToString()
        {
            if (Value != null)
                return Value.ToString();
            else
                return base.ToString();
        }
        public List<string> Errors { get; set; }
    }
}
