using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace Lava3.Core.Model
{
    public partial class CarMillageSummary
    {

        public CarMillageSummary(ExcelWorksheet sheet, Dictionary<string, ColumnHeader> ch)
        {
            int rownum = 4;
            this.DistanceStart = new DataTypes.ColumnDecimal(sheet, rownum, ch["Start"]);
            this.DistanceEnd =  new DataTypes.ColumnDecimal(sheet, rownum, ch["End"]);
            this.DistanceDaily = new DataTypes.ColumnDecimal(sheet, rownum, ch["Daily"]);
            this.DistanceWeekly = new DataTypes.ColumnDecimal(sheet, rownum, ch["Weekly"]);
            this.DistanceTotal = new DataTypes.ColumnDecimal(sheet, rownum, ch["Total"]);
            this.RebateWeekly = new DataTypes.ColumnDecimal(sheet, rownum, ch["Weekly Rebate"]);
            this.RebateTotal = new DataTypes.ColumnDecimal(sheet, rownum, ch["Total Rebate"]);
            this.TollCharges = new DataTypes.ColumnDecimal(sheet, rownum, ch["Toll Charges"]);

            this.RebatePerUnit = new DataTypes.ColumnDecimal(sheet, 2, ch["Total Rebate"]);
        }
        public override string ToString()
        {
            if(this.DistanceTotal!=null)
            {
                return $"({DistanceTotal} @ £{RebatePerUnit}) + Toll Charge £{TollCharges} = £{RebateTotal}";
            }
            else{
                return base.ToString();
            }
        }
    }
}
