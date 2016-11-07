using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Lava3.Core.DataTypes;

namespace Lava3.Core.Model
{
    public partial class CarMillageSummary
    {
        public ColumnDecimal DistanceStart { get; set; }
        public ColumnDecimal DistanceEnd { get; set; }
        public ColumnDecimal DistanceDaily { get; set; }
        public ColumnDecimal DistanceWeekly { get; set; }
        public ColumnDecimal DistanceTotal { get; set; }
        public ColumnDecimal RebateWeekly { get; set; }
        public ColumnDecimal RebateTotal { get; set; }
        public ColumnDecimal RebatePerUnit { get; set; }
        public ColumnDecimal TollCharges { get; set; }

    }
}
