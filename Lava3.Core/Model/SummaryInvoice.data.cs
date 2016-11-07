using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Lava3.Core.DataTypes;

namespace Lava3.Core.Model
{
   public partial class SummaryInvoice
    {
        public ColumnString Customer { get; set; }
        public ColumnString InvoiceName { get; set; }
        public Uri InvoiceNameHyperLink { get; set; }
        public ColumnDateTime InvoiceDate { get; set; }
        public ColumnDateTime InvoicePaid { get; set; }
        public ColumnInt HoursInvoiced { get; set; }
        public ColumnInt DaysInvoiced { get; set; }
        public ColumnDecimal TotalPaid { get; set; }
        public ColumnDecimal InvoiceAmount { get; set; }


        public ColumnDecimal DayRate { get; private set; }
        public ColumnInt DaysToPay { get; private set; }
    }
}
