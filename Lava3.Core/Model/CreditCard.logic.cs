using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.Text.RegularExpressions;
using Lava3.Core.DataTypes;
using Lava3.Core.Properties;

namespace Lava3.Core.Model
{
    public partial class CreditCard
    {
        public CreditCard(ExcelWorksheet sheet, Dictionary<string, ColumnHeader> ch, int rownum, IEnumerable<Category> categoryRows)
        {
            RowNumber = rownum;
            PaidDate = new ColumnDateTime(sheet, rownum, ch["Paid Date"]);
            StatementDate = new ColumnDateTime(sheet, rownum, ch["Statement Date"]);
            TransactionDate = new ColumnDateTime(sheet, rownum, ch["Transaction Date"]);
            TransactionDescription = new ColumnString(sheet, rownum, ch["Transaction Description"]);
            TransactionAmount = new ColumnDecimal(sheet, rownum, ch["Transaction Amount"]);
            VatContent = new ColumnDecimal(sheet, rownum, ch["Vat Content"]);
            Postage = new ColumnDecimal(sheet, rownum, ch["P & P"]);

            Category = new ColumnString(sheet, rownum, ch["Category"]);
            Notes = new ColumnString(sheet, rownum, ch["Notes"]);
            if (sheet.Cells[rownum, ch["Notes"].ColumnNumber].Hyperlink != null)
            {
                NotesHyperLink = sheet.Cells[rownum, ch["Notes"].ColumnNumber].Hyperlink;
            }
            if (categoryRows.Any())
            {
                Categorise(categoryRows);
            }
            if(PaidDate.Value!=null && StatementDate.Value==null)
            {
                StatementDate.Errors.Add("Statement date is missing.");
            }
        }

        public override string ToString()
        {
            return $"{RowNumber} {TransactionDate} {Category} {TransactionDescription} {TransactionAmount}";
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
        public void Categorise(IEnumerable<Category> categories)
        {
            //Only categorise if value is not already set.
            if (!string.IsNullOrEmpty(Category.Value)) return;
            //
            Category localCategory = null;
            IEnumerable<Category> c;
            c = categories.Where(w => w.Description.Value.Equals(TransactionDescription.Value,
                                                           StringComparison.CurrentCultureIgnoreCase));
            if (!c.Any())
            {
                c = categories.Where(w => w.Description.Value.Contains("*"));
                if (c.Any())
                {
                    foreach (Category item in categories.Where(w => !w.Description.Value
                                                                      .Equals(Resources.DontMap, StringComparison.CurrentCultureIgnoreCase)))
                    {
                        var regex = new Regex(item.RegEx.Value, RegexOptions.IgnoreCase);
                        if (regex.Match(TransactionDescription.Value).Length > 0)
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

            if (localCategory != null)
            {
                if (!localCategory.Description.Value.Equals(Resources.DontMap, StringComparison.CurrentCultureIgnoreCase))
                {
                    Category = Common.ReplaceIfEmpty(Category, localCategory.AccountingCategory);
                    //TransactionDescription = category.Description;
                    if (!string.IsNullOrEmpty(localCategory.Notes.Value))
                    {
                        Notes = Common.ReplaceIfEmpty(Notes, localCategory.Notes); ;
                        NotesHyperLink = localCategory.NotesHyperLink;
                    }

                }
            }
            if (string.IsNullOrEmpty(Category.Value))
            {
                Category.Errors.Add("Missing Category");
            }

        }
    }
}
