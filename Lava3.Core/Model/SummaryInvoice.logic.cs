using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Lava3.Core.DataTypes;
using OfficeOpenXml;

namespace Lava3.Core.Model
{
    public partial class SummaryInvoice
    {
        public int RowNumber { get; private set; }
        public SummaryInvoice(ExcelWorksheet sheet, Dictionary<string, ColumnHeader> ch, int rownum)
        {
            RowNumber = rownum;
            this.Customer = new ColumnString(sheet, rownum, ch["Customer"]);
            this.InvoiceName = new ColumnString(sheet, rownum, ch["Invoice"]);
            this.InvoiceDate = new ColumnDateTime(sheet, rownum, ch["Invoice Date"]);
            this.InvoicePaid = new ColumnDateTime(sheet, rownum, ch["Date Funds Recieved"]);
            this.HoursInvoiced = new ColumnInt(sheet, rownum, ch["Hours Invoiced"]);
            this.DaysInvoiced = new ColumnInt(sheet, rownum, ch["Days Invoiced"]);
            this.InvoiceAmount = new ColumnDecimal(sheet, rownum, ch["Invoice Amount"]);
            this.TotalPaid = new ColumnDecimal(sheet, rownum, ch["Total Paid"]);

            if (InvoiceDate.Value != null && InvoicePaid.Value != null)
            {
                DaysToPay = new ColumnInt(sheet, rownum, ch["Days to pay"]);
                DaysToPay.Value = (int)((DateTime)InvoicePaid.Value - (DateTime)InvoiceDate.Value).TotalDays;
            }
            if(DaysInvoiced.Value !=0 && InvoiceAmount.Value !=0 )
            {
                DayRate = new ColumnDecimal(sheet, rownum, ch["Day Rate"]);
                DayRate.Value = InvoiceAmount.Value / DaysInvoiced.Value;                
            }
            ExcelAddress invoiceRange = new ExcelAddress(rownum, InvoiceName.ColumnNumber, rownum, InvoiceName.ColumnNumber);

            if(sheet.Cells[invoiceRange.Address].Hyperlink !=null)
            {
                this.InvoiceNameHyperLink = sheet.Cells[invoiceRange.Address].Hyperlink;
            }
        }

        public override string ToString()
        {
            if(RowNumber==0)     return base.ToString();

            return $"[{RowNumber}] {InvoiceDate} {InvoiceName} {InvoiceAmount}";

        }


    }
}
