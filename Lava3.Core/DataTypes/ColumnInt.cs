using System.Collections.Generic;
using OfficeOpenXml;

namespace Lava3.Core.DataTypes
{
    public class ColumnInt : IColumDataType
    {
        public ColumnInt()
        {
            Errors = new List<string>();
        }

        public ColumnInt(ExcelWorksheet sheet, int rownum, dynamic categoryColumns)
        {
            Errors = new List<string>();
            ColumnNumber = categoryColumns.ColumnNumber;
            Value = Common.ReplaceNullOrEmptyInt(sheet.Cells[rownum, ColumnNumber].Value);
        }

        public int? Value { get; set; }
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
