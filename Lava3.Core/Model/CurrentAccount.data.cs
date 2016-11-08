using System;
using System.Collections.Generic;
using Lava3.Core.DataTypes;

namespace Lava3.Core.Model
{
    public partial class CurrentAccount
    {
        public int RowNumber { get; set; }
        public ColumnDateTime Date { get; set; }
        public ColumnString Description { get; set; }
        public ColumnDecimal Debit { get; set; }
        public ColumnDecimal Credit { get; set; }
        public ColumnDecimal Balence { get; set; }
        public ColumnDecimal MonthlyBalence { get; set; }
        public ColumnDecimal YearlyBalence { get; set; }
        public ColumnString Category { get; set; }
        public ColumnString Notes { get; set; }
        public Uri NotesHyperLink { get; set; }
        public bool IsMonthlySummary { get; set; }
        public bool IsDivider { get; set; }
        public bool IsStartingBalence { get; set; }
        public bool IsCreditCard { get; set; }
        public IEnumerable<CreditCard> CreditCardTransactions { get; set; }
    }
}
