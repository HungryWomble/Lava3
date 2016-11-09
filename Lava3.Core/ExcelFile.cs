using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Lava3.Core.Model;
using OfficeOpenXml;
using System.Xml.Linq;
using Lava3.Core.DataTypes;
using System.Drawing;
using OfficeOpenXml.Style.XmlAccess;
using System.Reflection;

namespace Lava3.Core
{
    public class ExcelFile : IDisposable
    {
        public ExcelFile()
        {

        }
        public ExcelFile(string filename)
        {
            OpenPackage(filename);
        }
        public static class eWorkSheetLabels
        {
            public const string CreditCard = "CreditCard";
            public const string CategoryLookup = "Category LookUp";
            public const string CurrentAccount = "HSBC";
            public const string AnnualSummary = "Annual Summary";
            public const string CarMilage = "Car Mileage";
        }
        public struct eDescriptionKeys
        {
            public const string StartingBalance = "Balance At start of year";
            public const string Totals = "Totals";
            public struct AnnualSummary
            {
                public const string Invoices = "SALES INVOICES";
                public const string Expenses = "PAYMENTS MADE BY DIRECTOR PRIVATELY/CASH";
                public const string Summary = "BANK ACCOUNT PAYMENTS";
            }

        }

        private string UnknonwnCategory = "??????";
        #region Kill
        public void KillAllExcel()
        {
            foreach (Process item in Process.GetProcesses())
            {
                if (item.ProcessName.ToLower().Contains("excel"))
                {
                    item.Kill();
                }
            }
        }
        public void KillFile(string path)
        {
            foreach (Process item in Process.GetProcesses())
            {

                if (item.ProcessName.ToLower().Contains("excel"))
                {
                    FileInfo fi = new FileInfo(path);
                    if (item.MainWindowTitle == $"Microsoft Excel - {fi.Name}")
                    {
                        item.Kill();
                    }
                }
            }
        }

        #endregion

        /// <summary>
        /// Open the file for the user
        /// </summary>
        public void ShowFile(string path)
        {
            if (!File.Exists(path))
            {
                throw new FileNotFoundException($"Excel file could not be found '{path}'.");
            }
            Process openFile = new Process();
            Process.Start(path);

        }

        #region properties
        public ExcelPackage Package { get; set; }

        private ExcelWorksheet _SheetCategories;
        private ExcelWorksheet _SheetCreditCard;
        private ExcelWorksheet _SheetCurrentAccount;
        private ExcelWorksheet _SheetCarMileage;
        private ExcelWorksheet _SheetAnnualSummary;
        public IEnumerable<Category> CategoryRows { get; set; }
        public IEnumerable<CreditCard> CreditCardRows { get; set; }
        public IEnumerable<CurrentAccount> CurrentAccountRows { get; set; }
        public IEnumerable<SummaryInvoice> Invoices { get; set; }
        public IEnumerable<SummaryExpense> Expenses { get; set; }
        public Dictionary<string, ColumnHeader> CategoryColumns { get; set; }
        public CarMillageSummary MileageSummary { get; set; }
        #endregion
        /// <summary>
        /// Open the excel Package
        /// </summary>
        /// <param name="path"></param>
        public void OpenPackage(string path)
        {
            if (Package == null)
            {
                KillAllExcel();
                Package = new ExcelPackage(new FileInfo(path));
                _SheetCategories = Package.Workbook.Worksheets[eWorkSheetLabels.CategoryLookup];
            }
        }


        public void LoadCarSummary()
        {
            _SheetCarMileage = Package.Workbook.Worksheets[eWorkSheetLabels.CarMilage];
            var columnHeaders = Common.GetColumnHeaders(_SheetCarMileage, 3);
            MileageSummary = new CarMillageSummary(_SheetCarMileage, columnHeaders);
        }
        #region Load...
        public void LoadCurrentAccount()
        {
            LoadAndUpdateCreditCard();
            _SheetCurrentAccount = Package.Workbook.Worksheets[eWorkSheetLabels.CurrentAccount];
            var columnHeaders = Common.GetColumnHeaders(_SheetCurrentAccount, 2);
            int rownum = 3;
            var Rows = new List<CurrentAccount>();
            while (rownum <= _SheetCurrentAccount.Dimension.Rows)
            {
                CurrentAccount row = new CurrentAccount(_SheetCurrentAccount, columnHeaders, rownum, CategoryRows, CreditCardRows);

                rownum++;
                if (string.IsNullOrEmpty(row.Description.Value)) continue;
                Rows.Add(row);
            }
            //Remove boundries and monthly totals
            for (int i = Rows.Count - 1; i >= 0; i--)
            {
                if (Rows[i].IsDivider || Rows[i].IsMonthlySummary)
                {
                    Rows.RemoveAt(i);
                }
            }


            // sort by transaction date
            Rows = Rows.OrderBy(o => o.Date.Value)
                                                   .ToList();
            //Set the monthly and annual running totals.
            var retval = new List<CurrentAccount>();
            int currentMonth = -1;
            int previousMonth = ((DateTime)Rows[1].Date.Value).Month;
            Decimal? MonthlyTotal = 0m;
            int NewRowNumber = 0;
            retval.Add(Rows[0]);
            decimal? MonthlyDebit = 0;
            decimal? MonthlyCredit = 0;
            for (int i = 1; i < Rows.Count; i++)
            {
                var previous = Rows[i - 1];
                var current = Rows[i];
                previousMonth = ((DateTime)previous.Date.Value).Month;
                currentMonth = ((DateTime)current.Date.Value).Month;
                NewRowNumber++;
                if (currentMonth != previousMonth)
                {
                    var monthSummary = new CurrentAccount()
                    {
                        IsMonthlySummary = true,
                        RowNumber = NewRowNumber,
                        Notes = new ColumnString() { ColumnNumber = current.Notes.ColumnNumber },
                        Debit = new ColumnDecimal()
                        {
                            ColumnNumber = current.Notes.ColumnNumber,
                            Value = MonthlyDebit
                        },
                        Credit = new ColumnDecimal()
                        {
                            ColumnNumber = current.Notes.ColumnNumber,
                            Value = MonthlyCredit
                        }
                    };
                    retval.Add(monthSummary);
                    //
                    NewRowNumber++;

                    retval.Add(new CurrentAccount()
                    {
                        IsDivider = true,
                        RowNumber = NewRowNumber,
                        Notes = new ColumnString() { ColumnNumber = current.Notes.ColumnNumber }
                    });
                    NewRowNumber++;

                    MonthlyTotal = 0m;
                    MonthlyDebit = 0;
                    MonthlyCredit = 0;
                }
                current.RowNumber = NewRowNumber;
                decimal? transactionBalence = current.Credit.Value + current.Debit.Value;
                MonthlyTotal += transactionBalence;


                current.YearlyBalence.Value = previous.YearlyBalence.Value
                                                            + transactionBalence;
                current.MonthlyBalence.Value = MonthlyTotal;

                MonthlyCredit = MonthlyCredit + current.Credit.Value;
                MonthlyDebit = MonthlyDebit + current.Debit.Value;
                //Validation
                if (current.YearlyBalence.Value != current.Balence.Value)
                {
                    string msg = $"Balance and Yearly Balence do not match {current.Balence} != {current.YearlyBalence}";
                    current.YearlyBalence.Errors.Add(msg);
                    current.Balence.Errors.Add(msg);
                }
                retval.Add(Rows[i]);
            }
            

            CurrentAccountRows = retval;
        }
        /// <summary>
        /// Load the credit card into memory
        /// </summary>
        public void LoadCreditCard()
        {
            LoadAndUpdateCategory();

            _SheetCreditCard = Package.Workbook.Worksheets[eWorkSheetLabels.CreditCard];

            var columnHeaders = Common.GetColumnHeaders(_SheetCreditCard, 1);
            var rows = new List<CreditCard>();
            int rownum = 1;
            while (rownum <= _SheetCreditCard.Dimension.Rows)
            {
                rownum++;
                CreditCard row = new CreditCard(_SheetCreditCard, columnHeaders, rownum, CategoryRows);
                if (row.TransactionDate?.Value == null) continue;
                rows.Add(row);
            }

            CreditCardRows = rows;
        }
        public void LoadAnnualSummary()
        {
            LoadCarSummary();
            LoadAndUpdateCurrentAccount();
            _SheetAnnualSummary = Package.Workbook.Worksheets[eWorkSheetLabels.AnnualSummary];
            var chExpences = Common.GetColumnHeaders(_SheetAnnualSummary, 1, eDescriptionKeys.AnnualSummary.Expenses, 2);
            var chInvoices = Common.GetColumnHeaders(_SheetAnnualSummary, 1, eDescriptionKeys.AnnualSummary.Invoices);


            #region Get the expences

            int ExpenseStartRownumber = Common.GetRownumberForKey(_SheetAnnualSummary, eDescriptionKeys.AnnualSummary.Expenses, 1) + 3;
            int ExpenseTotalRowNumer = Common.GetRownumberForKey(_SheetAnnualSummary, eDescriptionKeys.Totals, 2, ExpenseStartRownumber);
            List<SummaryExpense> localExpense = new List<SummaryExpense>();
            if (ExpenseTotalRowNumer - ExpenseStartRownumber != 0)
            {
                for (int rownum = ExpenseStartRownumber; rownum < ExpenseTotalRowNumer; rownum++)
                {
                    SummaryExpense expense = new SummaryExpense(_SheetAnnualSummary, chExpences, rownum);
                    if (!string.IsNullOrEmpty(expense.Description.Value))
                    {
                        localExpense.Add(expense);
                    }
                }
            }
            Expenses = localExpense;
            #endregion

            #region get the invoices
            int InvoiceStartRownumber = Common.GetRownumberForKey(_SheetAnnualSummary, eDescriptionKeys.AnnualSummary.Invoices, 1) + 2;
            int InvoiceTotalRowNumer = Common.GetRownumberForKey(_SheetAnnualSummary, eDescriptionKeys.Totals, 2, InvoiceStartRownumber + 1);
            List<SummaryInvoice> localInvoices = new List<SummaryInvoice>();
            if (InvoiceTotalRowNumer - InvoiceStartRownumber != 0)
            {
                for (int rownum = InvoiceStartRownumber; rownum < InvoiceTotalRowNumer; rownum++)
                {
                    localInvoices.Add(new SummaryInvoice(_SheetAnnualSummary, chInvoices, rownum));
                }
            }
            Invoices = localInvoices;
            #endregion
        }
        /// <summary>
        /// Load the category into Memory
        /// </summary>
        public void LoadCategory()
        {
            CategoryColumns = Common.GetColumnHeaders(_SheetCategories, 1);

            List<Category> accountingCategories = new List<Category>();
            int rownum = 1;
            while (rownum <= _SheetCategories.Dimension.Rows)
            {
                ColumnString description = new ColumnString(_SheetCategories, rownum, CategoryColumns["Description"]);
                if (!string.IsNullOrEmpty(description.Value))
                {
                    var row = new Category()
                    {
                        Description = description,
                        AccountingCategory = new ColumnString(_SheetCategories, rownum, CategoryColumns["Category"]),
                        Notes = new ColumnString(_SheetCategories, rownum, CategoryColumns["Notes"])
                    };
                    if (_SheetCategories.Cells[rownum, CategoryColumns["Notes"].ColumnNumber].Hyperlink != null)
                    {
                        row.NotesHyperLink = _SheetCategories.Cells[rownum, CategoryColumns["Notes"].ColumnNumber].Hyperlink;
                    }

                    accountingCategories.Add(row);
                }
                rownum++;
            }
            if (!accountingCategories.Any())
            {
                throw new IndexOutOfRangeException("No Categories could be found");
            }
            //Sort by description
            accountingCategories = accountingCategories.OrderBy(o => o.Description.Value).ToList();
            // Set the duplicate flags
            IEnumerable<string> duplicateDescriptions = accountingCategories
                                        .GroupBy(g => g.Description.Value)
                                        .Where(w => !string.IsNullOrEmpty(w.Key) && w.Count() > 1)
                                        .Select(s => s.Key);
            foreach (string key in duplicateDescriptions)
            {
                foreach (Category c in accountingCategories.Where(w => !string.IsNullOrEmpty(w.Description.Value) &&
                                                                       w.Description.Value.Equals(key, StringComparison.CurrentCultureIgnoreCase)))
                {
                    c.IsDuplicateDescription = true;
                }
            }
            IEnumerable<string> duplicateNotes = accountingCategories
                                       .GroupBy(g => g.Notes.Value)
                                       .Where(w => !string.IsNullOrEmpty(w.Key) && w.Count() > 1)
                                       .Select(s => s.Key);
            foreach (string key in duplicateNotes)
            {
                foreach (Category c in accountingCategories.Where(w => !string.IsNullOrEmpty(w.Notes.Value) &&
                                                                      w.Notes.Value.Equals(key, StringComparison.CurrentCultureIgnoreCase)))
                {
                    c.IsDuplicateNotes = true;
                }
            }

            CategoryRows = accountingCategories;
        }
        #endregion

        #region LoadAndUpdate

        public void LoadAndUpdateCategory()
        {
            LoadCategory();
            UpsertCatergory();
        }

        public void LoadAndUpdateCreditCard()
        {
            LoadCreditCard();
            UpsertCreditCard();
        }
        public void LoadAndUpdateAnnualSummary()
        {
            LoadAnnualSummary();
            UpsertAnnualSummary();
        }

        public void LoadAndUpdateCurrentAccount()
        {
            LoadCurrentAccount();
            UpsertCurrentAccount();
        }
        #endregion

        public void Save()
        {
            Package.Save();
        }

        public void SaveAndClose()
        {
            Save();
            ClosePackage();
        }


        public void ClosePackage()
        {
            if (Package != null)
            {
                Package.Dispose();
            }
            Package = null;

        }


        public void Dispose()
        {
            ClosePackage();
        }
        private void CreateStyleHyperLink(ExcelWorksheet sheet, string stylename)
        {
            try
            {
                ExcelNamedStyleXml styleHyperlink = sheet.Workbook.Styles.CreateNamedStyle(stylename);
                styleHyperlink.Style.Font.UnderLine = true;
                styleHyperlink.Style.Font.Color.SetColor(Color.Blue);
            }
            catch { }
        }

        /// <summary>
        /// 1. delete all rows in the category worksheet
        /// 2. Write rows from category list into category worksheet.
        /// </summary>
        private void UpsertCatergory()
        {

            string stylenameHyperlink = "HyperLink";
            CreateStyleHyperLink(_SheetCategories, stylenameHyperlink);

            int rownum = 1;
            Common.DeleteRows(_SheetCategories, 2);
            foreach (Category item in CategoryRows)
            {
                rownum++;
                Common.UpdateCellString(_SheetCategories, rownum, item.Description);
                Common.UpdateCellString(_SheetCategories, rownum, item.AccountingCategory);

                if (item.NotesHyperLink == null)
                {
                    Common.UpdateCellString(_SheetCategories, rownum, item.Notes);
                }
                else
                {
                    ExcelRange cell = _SheetCategories.Cells[rownum, item.Notes.ColumnNumber];
                    cell.Hyperlink = item.NotesHyperLink;
                    cell.StyleName = stylenameHyperlink;
                    cell.Value = item.Notes.Value;
                }

                if (item.IsDuplicateDescription)
                {
                    Common.SetComment(_SheetCategories, rownum, item.Description.ColumnNumber, "Duplicate description.", Common.Colours.DuplicateColour);
                }
                if (item.IsDuplicateNotes)
                {
                    Common.SetComment(_SheetCategories, rownum, item.Notes.ColumnNumber, "Duplicate notes.", Common.Colours.DuplicateColour);
                }
            }
        }
        private void UpsertCreditCard()
        {
            if (CreditCardRows == null) return;

            //Create styles
            string stylenameHyperlink = "HyperLink";
            CreateStyleHyperLink(_SheetCreditCard, stylenameHyperlink);

            Common.DeleteRows(_SheetCreditCard, 2);
            int rownum = 1;
            foreach (CreditCard item in CreditCardRows)
            {
                rownum++;
                Common.UpdateCellDate(_SheetCreditCard, rownum, item.PaidDate);
                Common.UpdateCellDate(_SheetCreditCard, rownum, item.StatementDate);
                Common.UpdateCellDate(_SheetCreditCard, rownum, item.TransactionDate);
                Common.UpdateCellString(_SheetCreditCard, rownum, item.TransactionDescription);
                Common.UpdateCellString(_SheetCreditCard, rownum, item.Category);
                Common.UpdateCellDecimal(_SheetCreditCard, rownum, item.TransactionAmount);
                Common.UpdateCellDecimal(_SheetCreditCard, rownum, item.VatContent);
                Common.UpdateCellDecimal(_SheetCreditCard, rownum, item.Postage);
                Common.UpdateHyperLink(_SheetCreditCard, rownum, item.Notes, item.NotesHyperLink, stylenameHyperlink);

            }
            //Create conditional formating
            int categoryColumn = CreditCardRows.First().Category.ColumnNumber;
            ExcelAddress categoryAddress = new ExcelAddress(2,
                                                            categoryColumn,
                                                            rownum - 1,
                                                            categoryColumn);

            var cf = _SheetCreditCard.ConditionalFormatting.AddContainsBlanks(categoryAddress);
            cf.Style.Fill.BackgroundColor.Color = Common.Colours.ErrorColour;
        }
        private void UpsertCurrentAccount()
        {
            if (CategoryRows == null ||
                CreditCardRows == null ||
                CurrentAccountRows == null) return;

            Common.DeleteRows(_SheetCurrentAccount, 3);
            //Create styles
            string stylenameHyperlink = "HyperLink";
            CreateStyleHyperLink(_SheetCurrentAccount, stylenameHyperlink);


            int rownum = 2;
            foreach (var item in CurrentAccountRows)
            {
                rownum++;
                string CategoryMissing = "Category missing";
                if (item.IsStartingBalence)
                {
                    Common.SetRowColour(_SheetCurrentAccount, rownum, item.Notes.ColumnNumber, Common.Colours.StartingBalance, true);
                    CategoryMissing = null;
                }
                else if (item.IsDivider || item.IsMonthlySummary)
                {
                    Common.SetRowColour(_SheetCurrentAccount, rownum, item.Notes.ColumnNumber, Common.Colours.DividerColour, true);
                    CategoryMissing = null;
                }
                Common.UpdateCellDate(_SheetCurrentAccount, rownum, item.Date);
                Common.UpdateCellString(_SheetCurrentAccount, rownum, item.Description);
                Common.UpdateCellDecimal(_SheetCurrentAccount, rownum, item.Debit);
                Common.UpdateCellDecimal(_SheetCurrentAccount, rownum, item.Credit);
                Common.UpdateCellDecimal(_SheetCurrentAccount, rownum, item.Balence);
                Common.UpdateCellDecimal(_SheetCurrentAccount, rownum, item.MonthlyBalence);
                Common.UpdateCellDecimal(_SheetCurrentAccount, rownum, item.YearlyBalence);
                Common.UpdateCellString(_SheetCurrentAccount, rownum, item.Category, CategoryMissing);
                Common.UpdateHyperLink(_SheetCurrentAccount, rownum, item.Notes, item.NotesHyperLink, stylenameHyperlink);

                //Create conditional formating
                int categoryColumn = CurrentAccountRows.Last().Category.ColumnNumber;
                ExcelAddress categoryAddress = new ExcelAddress(rownum,
                                                                categoryColumn,
                                                                rownum,
                                                                categoryColumn);


                //Wrap category text
                _SheetCurrentAccount.Cells[categoryAddress.Address].Style.WrapText = true;

            }
        }

        private void UpsertAnnualSummary()
        {
          
            string stylenameHyperlink = "HyperLink";
            SummaryExpense FirstExpense = Expenses.First();
            SummaryInvoice FirstInvoice = Invoices.FirstOrDefault();
            int LastExpenseColumnNumber = Common.GetLastColumnNumber(FirstExpense);
            int LastInvoiceColumnNumber = Common.GetLastColumnNumber(FirstInvoice);
            Dictionary<string, ColumnHeader> chExpences = Common.GetColumnHeaders(_SheetAnnualSummary, 1, eDescriptionKeys.AnnualSummary.Expenses, 2);
            Dictionary<string, ColumnHeader> chInvoices = Common.GetColumnHeaders(_SheetAnnualSummary, 1, eDescriptionKeys.AnnualSummary.Invoices);
            Dictionary<string, ColumnHeader> chSummary = new Dictionary<string, ColumnHeader>();

            int rownum = 4;
            Common.DeleteRows(_SheetAnnualSummary, 4);
            #region Add Summary


            var caRows = CurrentAccountRows.Where(w => !w.IsDivider &&
                                                       !w.IsMonthlySummary &&
                                                       !w.IsStartingBalence).ToList();

            caRows.Sort(delegate (CurrentAccount x, CurrentAccount y)
            {
                if (x.Description == null && y.Description == null) return 0;
                else if (x.Description == null) return -1;
                else if (y.Description == null) return 1;
                else return x.Description.Value.CompareTo(y.Description.Value);
            });
            
            //Build summary Headers
            int summaryColumnHeaderNumber = 0;
            List<string> SummaryHeaders = "Date,Payee,Reconsiliation,Total Spent,VAT,Dividends,Salary,Expenses,PAYE".Split(',').ToList();
            List<string> SummaryHeaders2 = new List<string>();
            foreach (CreditCard item in CreditCardRows.Where(w=>!SummaryHeaders.Any( a=>a.Equals(w.Category.Value,StringComparison.CurrentCultureIgnoreCase)) ))
            {
                if (item.Category == null || string.IsNullOrEmpty(item.Category.Value))
                {
                    if (!SummaryHeaders2.Contains(UnknonwnCategory))
                    {
                        SummaryHeaders2.Add(UnknonwnCategory);
                    }
                }
                else if (!SummaryHeaders2.Contains(item.Category.Value))
                {
                    SummaryHeaders2.Add(item.Category.Value);
                }
            }

            foreach (CurrentAccount item in caRows.Where(w => !SummaryHeaders.Any(a => a.Equals(w.Category.Value, StringComparison.CurrentCultureIgnoreCase))))
            {
                if (item.Description != null &&
                    item.Description.Value.Equals("COMMERCIAL CARD", StringComparison.CurrentCultureIgnoreCase))
                    continue;

                if (item.Category == null || string.IsNullOrEmpty(item.Category.Value))
                {
                    if (!SummaryHeaders2.Contains(UnknonwnCategory))
                    {
                        SummaryHeaders2.Add(UnknonwnCategory);
                    }
                    item.Category = new ColumnString() { Value = UnknonwnCategory };
                }
                else if (!SummaryHeaders2.Contains(item.Category.Value))
                {
                    SummaryHeaders2.Add(item.Category.Value);
                }
            }
            SummaryHeaders2.Sort();
            SummaryHeaders.AddRange(SummaryHeaders2);

            foreach (string item in SummaryHeaders)
            {
                summaryColumnHeaderNumber++;
                chSummary.Add(item.Trim(), new ColumnHeader() { Header = item.Trim(), ColumnNumber = summaryColumnHeaderNumber });
            }

            //add summary headers to sheet
            Common.SetHeaders(_SheetAnnualSummary, rownum, chSummary);

            //add summary
            
            int summaryColumnsCount = chSummary.Count();
            foreach (CurrentAccount currentAccount in caRows.Where(w => !w.IsDivider && 
                                                                                    !w.IsMonthlySummary && 
                                                                                    !w.IsStartingBalence))
            {
                //Only processing the debits here
                if (currentAccount.Debit.Value == 0)
                    continue;
                rownum++;
              
                Common.UpdateCellDate(_SheetAnnualSummary, rownum, new ColumnDateTime() { ColumnNumber = 1, Value = currentAccount.Date.Value });
                Common.UpdateCellString(_SheetAnnualSummary, rownum, new ColumnString() { ColumnNumber = 2, Value = currentAccount.Description.Value });
                Common.AddFormulaDecimal(_SheetAnnualSummary, rownum, 3, $"D{rownum}-{currentAccount.Debit.Value}");

                Common.AddSumFormula(_SheetAnnualSummary, rownum, 4, rownum, 5, rownum, summaryColumnsCount);
                if (!currentAccount.IsCreditCard)
                {
                    int colnum = chSummary.Single(w => w.Key.Equals(currentAccount.Category.Value, StringComparison.CurrentCultureIgnoreCase)).Value.ColumnNumber;
                    Common.UpdateCellDecimal(_SheetAnnualSummary, rownum, new ColumnDecimal() { ColumnNumber = colnum, Value = currentAccount.Debit.Value });
                }
                else
                {
                    foreach (CreditCard creditCard in currentAccount.CreditCardTransactions)
                    {
                        if (creditCard.Category == null || string.IsNullOrEmpty(creditCard.Category.Value))
                        {
                            creditCard.Category = new ColumnString() { Value = UnknonwnCategory };
                        }

                        int colnum = chSummary.Single(w => w.Key.Equals(creditCard.Category.Value, StringComparison.CurrentCultureIgnoreCase)).Value.ColumnNumber;
                        Common.UpdateCellDecimal(_SheetAnnualSummary, rownum, new ColumnDecimal() { ColumnNumber = colnum, Value = creditCard.TransactionAmount.Value });
                    }
                }


            }
           
            //add summary Totals
            rownum++;
            for (int i = 3; i <= chSummary.Count(); i++)
            {
                Common.SetTotal(_SheetAnnualSummary, rownum, 5, i);
            }
            Common.SetRowColour(_SheetAnnualSummary, rownum, chSummary.Count(), Common.Colours.TotalsColour, true);
            #endregion

            #region Add expenses
            rownum += 3;
            //Add header
            Common.UpdateCellString(_SheetAnnualSummary, rownum,
                                    new ColumnString() { ColumnNumber = 1, Value = eDescriptionKeys.AnnualSummary.Expenses },
                                    "",
                                    true);
            //////
            rownum++;
            Common.UpdateCellString(_SheetAnnualSummary, rownum,
                                    new ColumnString()
                                    {
                                        ColumnNumber = Expenses.First().VAT.ColumnNumber,
                                        Value = "If applicable"
                                    },
                                    "",
                                    false);


            Common.SetRowColour(_SheetAnnualSummary, rownum, LastExpenseColumnNumber, Common.Colours.HeaderColour, true);
            //////
            rownum++;
            int firstExpenseRow = rownum;
            Common.SetHeaders(_SheetAnnualSummary, rownum, chExpences, FirstExpense);

            foreach (SummaryExpense expense in Expenses)
            {
                rownum++;
                Common.UpdateCellDate(_SheetAnnualSummary, rownum, expense.Date);
                Common.UpdateCellString(_SheetAnnualSummary, rownum, expense.Description);
                Common.UpdateCellDecimal(_SheetAnnualSummary, rownum, expense.Value);
                Common.UpdateCellDecimal(_SheetAnnualSummary, rownum, expense.VAT);
                var SumAddress = new ExcelAddress(rownum, 3, rownum, 3);
                var SumRange = new ExcelAddress(rownum, 4, rownum, chExpences.Count());
                _SheetAnnualSummary.Cells[SumAddress.Address].Formula = $"SUM({SumRange.Address})";
            }
            rownum += 2;

            //Expenses Total row
            _SheetAnnualSummary.Cells[rownum, FirstExpense.Description.ColumnNumber].Value = eDescriptionKeys.Totals;
            for (int i = 3; i <= chExpences.Count(); i++)
            {
                Common.SetTotal(_SheetAnnualSummary, rownum, firstExpenseRow, i);
            }
            ExcelAddress TotalOwedAddress = new ExcelAddress(rownum, 4, rownum, 4);
            ExcelAddress TotalExpenseAddress = new ExcelAddress(rownum, 3, rownum, 3);
            string totalOwedFormula = $"{TotalExpenseAddress}+{_SheetAnnualSummary.Cells[TotalOwedAddress.Address].Formula}";
            _SheetAnnualSummary.Cells[TotalOwedAddress.Address].Formula = totalOwedFormula;

            Common.SetRowColour(_SheetAnnualSummary, rownum, LastExpenseColumnNumber, Common.Colours.TotalsColour, true);
            rownum += 2;
            #endregion

            #region Add Invoices

            Common.UpdateCellString(_SheetAnnualSummary, rownum,
                                    new ColumnString() { ColumnNumber = 1, Value = eDescriptionKeys.AnnualSummary.Invoices },
                                    "",
                                    true);
            rownum++;
            //Set the expense header
            Common.SetHeaders(_SheetAnnualSummary, rownum, chInvoices, FirstInvoice);


            //set the expense data
            int FirstInvoiceRow = rownum + 1;
            foreach (SummaryInvoice invoice in Invoices)
            {
                rownum++;
                Common.UpdateCellString(_SheetAnnualSummary, rownum, invoice.Customer);
                Common.UpdateHyperLink(_SheetAnnualSummary, rownum, invoice.InvoiceName, invoice.InvoiceNameHyperLink, stylenameHyperlink);
                Common.UpdateCellDate(_SheetAnnualSummary, rownum, invoice.InvoiceDate);
                Common.UpdateCellDate(_SheetAnnualSummary, rownum, invoice.InvoicePaid);
                Common.UpdateCellInt(_SheetAnnualSummary, rownum, invoice.DaysToPay);
                Common.UpdateCellInt(_SheetAnnualSummary, rownum, invoice.HoursInvoiced);
                Common.UpdateCellInt(_SheetAnnualSummary, rownum, invoice.DaysInvoiced);
                Common.UpdateCellDecimal(_SheetAnnualSummary, rownum, invoice.InvoiceAmount);
                Common.UpdateCellDecimal(_SheetAnnualSummary, rownum, invoice.TotalPaid);
                Common.UpdateCellDecimal(_SheetAnnualSummary, rownum, invoice.DayRate);
            }
            rownum += 2;
            //Invoice Total Row
            _SheetAnnualSummary.Cells[rownum, FirstInvoice.Customer.ColumnNumber].Value = eDescriptionKeys.Totals;
            Common.SetTotal(_SheetAnnualSummary, rownum, FirstInvoiceRow, FirstInvoice.HoursInvoiced.ColumnNumber);
            Common.SetTotal(_SheetAnnualSummary, rownum, FirstInvoiceRow, FirstInvoice.DaysInvoiced.ColumnNumber);
            Common.SetTotal(_SheetAnnualSummary, rownum, FirstInvoiceRow, FirstInvoice.InvoiceAmount.ColumnNumber);
            Common.SetTotal(_SheetAnnualSummary, rownum, FirstInvoiceRow, FirstInvoice.TotalPaid.ColumnNumber);
            Common.SetRowColour(_SheetAnnualSummary, rownum, LastInvoiceColumnNumber, Common.Colours.TotalsColour, true);
            #endregion

        }
    }
}
