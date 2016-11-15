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


        public ColumnString Description { get; set; }
        public ColumnDecimal Value { get;set;}
        public ColumnDateTime Date { get; set; }
        public ColumnDecimal VAT { get; set; }
        public ColumnString Category { get; set; }
    }
}
