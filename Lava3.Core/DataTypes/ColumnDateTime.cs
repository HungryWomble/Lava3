using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace Lava3.Core.DataTypes
{
    public class ColumnDateTime: IColumDataType
    {
        public ColumnDateTime()
        {
            Errors = new List<string>();
        }

        public ColumnDateTime(ExcelWorksheet sheet, int rownum, dynamic categoryColumns)
        {
            Errors = new List<string>();
            ColumnNumber = categoryColumns.ColumnNumber;
            Value = Common.ReplaceNullOrEmptyDateTime(sheet.Cells[rownum, ColumnNumber].Value);
        }

        public DateTime? Value { get; set; }
        public int ColumnNumber { get; set; }

        public override string ToString()
        {
            if(Value!=null)
                return Value.ToString();
            else
                return base.ToString();
        }
        public List<string> Errors { get; set; }
    }
  
}
