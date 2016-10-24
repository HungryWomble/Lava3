using System;
using Lava3.Core.DataTypes;

namespace Lava3.Core.Model
{
    public partial class CreditCard
    {

        public int RowNumber { get; set; }
        public ColumnDateTime PaidDate { get; set; }
        public ColumnDateTime StatementDate { get; set; }
        public ColumnDateTime TransactionDate { get; set; }
        public ColumnDecimal TransactionAmount { get; set; }
        public ColumnString TransactionDescription { get; set; }
        public ColumnString Category { get; set; }
        public ColumnString CategoryError { get; set; }
        public ColumnDecimal VatContent { get; set; }
        public ColumnDecimal Postage { get; set; }
        public ColumnString Notes { get; set; }
        public Uri NotesHyperLink { get; set; }
        
    }
}
