using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lava3.Core
{
    public static class Extensions
    {
        /// <summary>
        /// Get the first day of the month.
        /// </summary>
        public static DateTime? FirstDayOfMonth(this DateTime date)
        {
            if (date == null) return null;

            return new DateTime(date.Year, date.Month, 1);

        }
        /// <summary>
        /// Get the first day of the month.
        /// </summary>
        public static DateTime? LastDayOfPreviousMonth(this DateTime date)
        {
            if (date == null) return null;

            return new DateTime(date.Year, date.Month, 1).AddDays(-1);

        }
    }
}
