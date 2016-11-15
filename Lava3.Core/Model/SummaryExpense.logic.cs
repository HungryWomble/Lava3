using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Lava3.Core.DataTypes;
using OfficeOpenXml;

namespace Lava3.Core.Model
{
    public partial class SummaryExpense
    {

        public int RowNumber { get; private set; }
        public SummaryExpense(ExcelWorksheet sheet, Dictionary<string, ColumnHeader> ch, int rownum)
        {
            RowNumber = rownum;
            this.Date = new ColumnDateTime(sheet, rownum, ch["Date"]);
            this.Description = new ColumnString(sheet, rownum, ch["Description"]);
            this.VAT = new ColumnDecimal(sheet, rownum, ch["V.A.T."]);

            foreach (var header in ch)
            {
                int colnum = header.Value.ColumnNumber;
                if (colnum <= VAT.ColumnNumber &&
                    colnum != ch["Total Owed To Director"].ColumnNumber)
                    continue;
                ExcelRange cell = sheet.Cells[rownum, colnum];

                if (cell.Value != null &&
                    Common.ReplaceNullOrEmptyDecimal(cell.Value) != 0)
                {
                    this.Category = new ColumnString()
                    {
                        ColumnNumber = colnum,
                        Value = header.Value.Header
                    };
                    this.Value = new ColumnDecimal()
                    {
                        ColumnNumber = colnum,
                        Value = Common.ReplaceNullOrEmptyDecimal(cell.Value)
                    };
                    break;
                }
            }
            //if (this.IsExpenseRefund)
            //{
            //    this.Value = ExcelRange cell = sheet.Cells[rownum, colnum];
            //}


        }

        public override string ToString()
        {
            if (Description == null)
                return base.ToString();

            return $"{Date} | {Description.Value.PadRight(25).Substring(0, 25)} |{Category}";

        }
        /// <summary>
        /// Is this expense a refund
        /// </summary>
        public bool IsExpenseRefund
        {
            get
            {
                return Description.Value.ToLower().Contains("expense paid") ||
                       Description.Value.ToLower().Contains("expenses paid");
            }
        }
    }
}
