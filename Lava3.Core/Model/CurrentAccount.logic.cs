using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Lava3.Core.DataTypes;
using OfficeOpenXml;

namespace Lava3.Core.Model
{
    public partial class CurrentAccount
    {
        public override string ToString()
        {
            if (RowNumber == 0)
            {
                return base.ToString();
            }
            else if (IsStartingBalence)
            {
                return $"{RowNumber} Starting Balence";
            }
            else if (IsMonthlySummary)
            {
                return $"{RowNumber} ------ Monthly Summary";
            }
            else if (IsDivider)
            {
                return $"{RowNumber} ----------------------------";
            }
            else
            {
                return $"{RowNumber} {Date} {Category} {Description} {MonthlyBalence}";
            }
        }


        public CurrentAccount( ExcelWorksheet sheet, Dictionary<string, dynamic> ch, int rownum, IEnumerable<Category> categories, IEnumerable<CreditCard> ccRows)
        {
            RowNumber = rownum;
            Date = Common.ReplaceNullOrEmptyDateTime(sheet.Cells[rownum, ch["Date"].ColumnNumber].Value);
            Description = new ColumnString( sheet, rownum, ch["Description"]);
            Debit = new ColumnDecimal( sheet, rownum, ch["Debit"]);
            Credit = new ColumnDecimal( sheet, rownum, ch["Credit"]);
            Balence = new ColumnDecimal( sheet, rownum, ch["Balence"]);
            MonthlyBalence = new ColumnDecimal( sheet, rownum, ch["Monthly"]);
            YearlyBalence = new ColumnDecimal( sheet, rownum, ch["Yearly"]);
            Category = new ColumnString( sheet, rownum, ch["Category Override"]);
            Notes = new ColumnString( sheet, rownum, ch["Notes"]);
            if (Date == null)
            {
                IsDivider = true;
                if (Debit != null || Credit != null)
                {
                    IsMonthlySummary = true;
                }
            }

            IsStartingBalence = (rownum == 3);

            // set the categories
            Categorise(categories, ccRows);
        }

        /// <summary>
        /// Find categories in the following order
        /// 1/ Direct Match
        /// 2. WildCard at start
        /// 3. WildCard at end
        /// 4. WildCard at end and start
        /// </summary>
        /// <param name="row"></param>
        /// <returns></returns>
        public void Categorise(IEnumerable<Category> categories, IEnumerable<CreditCard> ccRows)
        {
            if (string.IsNullOrEmpty(Description.Value)) return;
            //
            Category category = null;
            IEnumerable<Category> c;
            c = categories.Where(w => w.Description.Value.Equals(Description.Value,
                                                           StringComparison.CurrentCultureIgnoreCase));
            if (!c.Any())
            {
                c = categories.Where(w => w.Description.Value.Contains("*"));
                if (c.Any())
                {
                    foreach (Category item in categories.Where(w => !w.Description.Value
                                                                      .Equals("dont map", StringComparison.CurrentCultureIgnoreCase)))
                    {
                        var regex = new Regex(item.RegEx.Value, RegexOptions.IgnoreCase);
                        if (regex.Match(Description.Value).Length > 0)
                        {
                            category = item;
                            break;
                        }
                    }
                }
            }
            else
            {
                category = c.Single();
            }
            if (category!=null &&
                    ccRows != null &&
                    category.AccountingCategory.Value.Equals("CC:HSBC", StringComparison.CurrentCultureIgnoreCase))
            {
                IEnumerable<CreditCard> paid = ccRows.Where(w => w.PaidDate.Value == this.Date);
                decimal? paidTotal = 0;
                if (paid == null || !paid.Any())
                {
                    Category.Errors.Add("Can not find any credit card transaction for a payment on this date.");
                }
                else if (paid != null && paid.Any())
                {
                    paidTotal = paid.Sum(s => s.TransactionAmount.Value);
                    if (paidTotal != this.Debit.Value)
                    {
                        Debit.Errors.Add("The debit value does not match the sum of the associated credit card purchases.");
                    }
                }
            }
            else if (category != null &&
                    !category.Description.Value.Equals("Dont Map", StringComparison.CurrentCultureIgnoreCase))
            {
                Category.Value = Common.ReplaceIfEmpty(Category.Value, category.AccountingCategory.Value);
                if (!string.IsNullOrEmpty(Notes.Value))
                {
                    Notes = category.Notes;
                    NotesHyperLink = category.NotesHyperLink;
                }
            }


        }
    }
}
