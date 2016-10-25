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


        public CurrentAccount(ExcelWorksheet sheet, Dictionary<string, dynamic> ch, int rownum, IEnumerable<Category> categories, IEnumerable<CreditCard> ccRows)
        {
            RowNumber = rownum;
            Date = new ColumnDateTime(sheet, rownum, ch["Date"]);
            Description = new ColumnString(sheet, rownum, ch["Description"]);
            Debit = new ColumnDecimal(sheet, rownum, ch["Debit"]);
            Credit = new ColumnDecimal(sheet, rownum, ch["Credit"]);
            Balence = new ColumnDecimal(sheet, rownum, ch["Balence"]);
            MonthlyBalence = new ColumnDecimal(sheet, rownum, ch["Monthly"]);
            YearlyBalence = new ColumnDecimal(sheet, rownum, ch["Yearly"]);
            Category = new ColumnString(sheet, rownum, ch["Category Override"]);
            Notes = new ColumnString(sheet, rownum, ch["Notes"]);
            if (sheet.Cells[rownum, ch["Notes"].ColumnNumber].Hyperlink != null)
            {
                NotesHyperLink = sheet.Cells[rownum, ch["Notes"].ColumnNumber].Hyperlink.OriginalUri;
            }
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
            if (string.IsNullOrEmpty(Description?.Value)
                || Description.Value.Equals(ExcelFile.eDescriptionKeys.StartingBalance)) return;
            //
            Category localCategory = null;
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
                            localCategory = item;
                            break;
                        }
                    }
                }
            }
            else
            {
                localCategory = c.Single();
            }
            if (localCategory != null &&
                    ccRows != null &&
                    localCategory.AccountingCategory.Value.Equals("CC:HSBC", StringComparison.CurrentCultureIgnoreCase))
            {
                IEnumerable<CreditCard> ccTransactions = ccRows.Where(w => w.PaidDate.Value == this.Date.Value);
                decimal? paidTotal = 0;
                if (ccTransactions == null || !ccTransactions.Any())
                {
                    Category.Errors.Add("Can not find any credit card transaction for a payment on this date.");
                }
                else if (ccTransactions != null && ccTransactions.Any())
                {
                    paidTotal = ccTransactions.Sum(s => s.TransactionAmount.Value);
                    if (paidTotal + this.Debit.Value != 0)
                    {
                        string errorMessage = "The value debited and the sum of the transactions in the catagory do not match.";
                        Debit.Errors.Add(errorMessage);
                        Category.Errors.Add(errorMessage);
                    }
                    StringBuilder sb = new StringBuilder();
                    bool HasNoCategory=false;
                    foreach (CreditCard item in ccTransactions)
                    {
                        if(string.IsNullOrEmpty(item.Category.Value))
                        {
                            HasNoCategory = true;
                        }
                        sb.AppendLine($" {item.TransactionAmount.Value:N2}, {item.Category.Value}");
                    }
                    Category.Value = sb.ToString().TrimEnd('\r', '\n');
                    if (HasNoCategory)
                    {
                        Category.Errors.Add("One or more transactions have not been categorised.");
                    }
                }
            }
            else if (localCategory != null &&
                    !localCategory.Description.Value.Equals("Dont Map", StringComparison.CurrentCultureIgnoreCase))
            {
                Category = Common.ReplaceIfEmpty(Category, localCategory.AccountingCategory);
                //TransactionDescription = category.Description;
                if (!string.IsNullOrEmpty(localCategory.Notes.Value))
                {
                    Notes = Common.ReplaceIfEmpty(Notes, localCategory.Notes); ;
                    NotesHyperLink = localCategory.NotesHyperLink;
                }
            }
            if (string.IsNullOrEmpty(Category.Value))
            {
                Category.Errors.Add("Missing Category");
            }


        }
    }
}
