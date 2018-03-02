using System.Collections.Generic;
using OfficeOpenXml;

namespace Lava3.Core.DataTypes
{
    public class ColumnDecimal : BaseColumn, IColumDataType
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

        public override string ToString()
        {
            if (Value != null)
                return Value.ToString();
            else
                return base.ToString();
        }
       
    }
    public abstract class BaseColumn
    {
        public List<string> Errors { get; set; }
        public string ColumnLetter { get { return Common.GetExcelColumnLetter(this.ColumnNumber); } }
        public int ColumnNumber { get; set; }
        /// <summary>
        /// The column code e.g. A1 or D4
        /// </summary>
        /// <param name="rownum"></param>
        /// <returns></returns>
        public string ColumnCode(int rownum)
        {
            return $"{ColumnLetter}{rownum}";
        }
    }
}
