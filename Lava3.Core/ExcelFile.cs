using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using Lava3.Core.Model;
using OfficeOpenXml;
using Lava3.Core.DataTypes;
using System.Drawing;
using OfficeOpenXml.Style.XmlAccess;
using Lava3.Core.Properties;

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
            public static string CreditCard = Resources.WorkSheelLabel_CreditCard;
            public static string CategoryLookup = Resources.WorkSheetLabel_CategoryLookup;
            public static string CurrentAccount = Resources.WorksheetLabel_CurrentAccount;
            public static string AnnualSummary = Resources.WorkSheetLabel_AnnualSummary;
            public static string CarMilage = Resources.WorkSheetLabel_CarMilage;
        }
        public struct eDescriptionKeys
        {
            public static string StartingBalance = Resources.DescriptionKey_BalanceAtStartOfYear;
            public static string Totals = Resources.DescriptionKey_Totals;
            public struct AnnualSummary
            {
                public static string Invoices = Resources.DescriptionKey_SalesInvoices;
                public static string Expenses = Resources.DescriptionKey_Expenses;
                public static string Summary = Resources.DescriptionKey_Summary;
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
                throw new FileNotFoundException($"{Resources.Error_ExcelFileCouldNotBeFound} '{path}'.");
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
            if (!File.Exists(path))
            {
                throw new FileNotFoundException(path);
            }

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
                if (!row.IsDivider && !row.IsMonthlySummary)
                {
                    Rows.Add(row);
                }
            }
            DateTime? endOfPreviousYear = Convert.ToDateTime(Rows.Where(f => !f.IsStartingBalence).OrderBy(o => o.Date.Value).First().Date.Value).LastDayOfPreviousMonth();
            Rows.Single(s => s.IsStartingBalence).Date.Value = endOfPreviousYear;
            // sort by transaction date
            Rows = Rows
                    .OrderBy(o => o.Date.Value)
                    .ThenBy(t => t.RowNumber)
                    .ToList();
            // Add dummy row to the end of rows so we get month end.

            // Set the monthly and annual running totals.
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
                    AddCurrentAccountMonthDivider(retval, ref NewRowNumber, ref MonthlyDebit, ref MonthlyCredit, columnHeaders, MonthlyTotal);
                }
                current.RowNumber = NewRowNumber;
                decimal? transactionBalence = Sum(current.Credit, current.Debit);
                MonthlyTotal += transactionBalence;


                current.CalculatedYearlyBalence.Value = previous.CalculatedYearlyBalence.Value
                                                            + transactionBalence;
                current.CalculatedMonthlyBalence.Value = MonthlyTotal;

                MonthlyCredit = MonthlyCredit + current.Credit.Value;
                MonthlyDebit = MonthlyDebit + current.Debit.Value;
                //Validation
                if (current.CalculatedYearlyBalence.Value != current.Balence.Value)
                {
                    string msg = $"{Resources.Validation_BalancesNotMatch} {current.Balence} != {current.CalculatedYearlyBalence}";
                    current.CalculatedYearlyBalence.Errors.Add(msg);
                    current.Balence.Errors.Add(msg);
                }
                retval.Add(Rows[i]);
            }
            AddCurrentAccountMonthDivider(retval, ref NewRowNumber, ref MonthlyDebit, ref MonthlyCredit, columnHeaders, MonthlyTotal);

            CurrentAccountRows = retval;
        }

        /// <summary>
        /// Sum up credit and debit
        /// </summary>
        /// <param name="value1"></param>
        /// <param name="value2"></param>
        /// <returns></returns>
        private decimal? Sum(ColumnDecimal value1, ColumnDecimal value2)
        {
            if (value1.Value == null) value1.Value = 0;
            if (value2.Value == null) value2.Value = 0;
            return value1.Value + value2.Value;
        }
        private static void AddCurrentAccountMonthDivider(List<CurrentAccount> retval,
                                                        ref int NewRowNumber,
                                                        ref decimal? MonthlyDebit,
                                                        ref decimal? MonthlyCredit,
                                                        Dictionary<string, ColumnHeader> ch,
                                                        decimal? MonthlyTotal)
        {

            var monthSummary = new CurrentAccount()
            {
                IsMonthlySummary = true,
                IsDivider = false,
                RowNumber = NewRowNumber,
                Notes = new ColumnString() { ColumnNumber = ch[Resources.ColumnHeader_Notes].ColumnNumber },
                Debit = new ColumnDecimal() { ColumnNumber = ch[Resources.ColumnHeader_Debit].ColumnNumber, Value = 0 },
                Credit = new ColumnDecimal() { ColumnNumber = ch[Resources.ColumnHeader_Credit].ColumnNumber, Value = 0 }
            };

            retval.Add(monthSummary);
            //
            NewRowNumber++;
            retval.Add(new CurrentAccount()
            {
                IsMonthlySummary = false,
                IsDivider = true,
                RowNumber = NewRowNumber
            });
            NewRowNumber++;

            MonthlyTotal = 0m;
            MonthlyDebit = 0;
            MonthlyCredit = 0;
        }

        /// <summary>
        /// Load the credit card into memory
        /// </summary>
        public void LoadCreditCard()
        {
            LoadAndUpdateCategory();

            _SheetCreditCard = Package.Workbook.Worksheets[eWorkSheetLabels.CreditCard];

            var columnHeaders = Common.GetColumnHeaders(_SheetCreditCard, 1);
            List<CreditCard> retval = new List<CreditCard>();
            int rownum = 1;
            while (rownum <= _SheetCreditCard.Dimension.Rows)
            {
                rownum++;
                CreditCard row = new CreditCard(_SheetCreditCard, columnHeaders, rownum, CategoryRows);
                if (row.TransactionDate?.Value == null) continue;
                retval.Add(row);
            }

            CreditCardRows = from r in retval
                             orderby r.StatementDate.Value
                             orderby r.TransactionDate.Value
                             select r;
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

            _SheetCategories = Package.Workbook.Worksheets[eWorkSheetLabels.CategoryLookup];
            CategoryColumns = Common.GetColumnHeaders(_SheetCategories, 1);

            List<Category> accountingCategories = new List<Category>();
            int rownum = 2;
            while (rownum <= _SheetCategories.Dimension.Rows)
            {
                ColumnString description = new ColumnString(_SheetCategories, rownum, CategoryColumns[Resources.ColumnHeader_Description]);
                if (!string.IsNullOrEmpty(description.Value))
                {
                    var row = new Category()
                    {
                        Description = description,
                        AccountingCategory = new ColumnString(_SheetCategories, rownum, CategoryColumns[Resources.ColumnHeader_Category]),
                        Notes = new ColumnString(_SheetCategories, rownum, CategoryColumns[Resources.ColumnHeader_Notes])
                    };
                    if (_SheetCategories.Cells[rownum, CategoryColumns[Resources.ColumnHeader_Notes].ColumnNumber].Hyperlink != null)
                    {
                        row.NotesHyperLink = _SheetCategories.Cells[rownum, CategoryColumns[Resources.ColumnHeader_Notes].ColumnNumber].Hyperlink;
                    }

                    accountingCategories.Add(row);
                }
                rownum++;
            }
            if (!accountingCategories.Any())
            {
                throw new IndexOutOfRangeException(Resources.Error_NoCategoriesCouldBeFound);
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
            accountingCategories.Sort(delegate (Category x, Category y)
            {
                if (x.Description == null && y.Description == null) return 0;
                else if (x.Description == null) return -1;
                else if (y.Description == null) return 1;
                else return x.Description.Value.CompareTo(y.Description.Value);
            });

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
                    //check file location
                    string path = Path.GetFullPath(Path.Combine(Package.File.Directory.FullName, cell.Hyperlink.OriginalString));
                    if (!File.Exists(path))
                    {
                        Common.SetComment(_SheetCategories, rownum, item.Notes.ColumnNumber, Resources.Validation_CannotResolveHyperLink, Common.Colours.ErrorColour);
                    }
                }

                if (item.IsDuplicateDescription)
                {
                    Common.SetComment(_SheetCategories, rownum, item.Description.ColumnNumber, Resources.Validation_DuplicateDescription, Common.Colours.DuplicateColour);
                }
                if (item.IsDuplicateNotes)
                {
                    Common.SetComment(_SheetCategories, rownum, item.Notes.ColumnNumber,Resources.Validation_DuplicateNotes, Common.Colours.DuplicateColour);
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
                Common.UpdateHyperLink(_SheetCreditCard, rownum, item.Notes, item.NotesHyperLink, stylenameHyperlink, Package.File.DirectoryName);

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

            Common.DeleteRows(_SheetCurrentAccount, 4);
            //Create styles
            string stylenameHyperlink = "HyperLink";
            CreateStyleHyperLink(_SheetCurrentAccount, stylenameHyperlink);


            int rownum = 2;
            int rowMonthStart = 0;
            int colDebit = CurrentAccountRows.First(f => !f.IsMonthlySummary && !f.IsDivider).Debit.ColumnNumber;
            int colCredit = CurrentAccountRows.First(f => !f.IsMonthlySummary && !f.IsDivider).Credit.ColumnNumber;
            int lastColumnNumber = CurrentAccountRows.First().Notes.ColumnNumber;
            int CategoryColumnNumber = CurrentAccountRows.First().Category.ColumnNumber;
            int yearlyBalenceOffset = -3;
            foreach (CurrentAccount item in CurrentAccountRows)
            {
                rownum++;
                string CategoryMissing = Resources.Validation_CategoryMissing;
                if (item.IsStartingBalence)
                {
                    Common.SetRowColour(_SheetCurrentAccount, rownum, lastColumnNumber, Common.Colours.StartingBalance, true);
                    CategoryMissing = null;
                    rowMonthStart = rownum + 1;
                }
                else if (item.IsDivider || item.IsMonthlySummary)
                {
                    Common.SetRowColour(_SheetCurrentAccount, rownum, lastColumnNumber, Common.Colours.DividerColour, true);
                    CategoryMissing = null;
                    item.Category = null;
                }
                else
                {
                    Common.UpdateCellDate(_SheetCurrentAccount, rownum, item.Date);
                    Common.UpdateCellString(_SheetCurrentAccount, rownum, item.Description);
                    Common.UpdateCellDecimal(_SheetCurrentAccount, rownum, item.Debit);
                    Common.UpdateCellDecimal(_SheetCurrentAccount, rownum, item.Credit);
                    Common.UpdateCellDecimal(_SheetCurrentAccount, rownum, item.Balence);
                    //Common.UpdateCellDecimal(_SheetCurrentAccount, rownum, item.MonthlyBalence);
                    Common.AddFormulaDecimal(_SheetCurrentAccount, rownum, item.CalculatedMonthlyBalence.ColumnNumber,
                        $"={item.CalculatedMonthlyBalence.ColumnCode(rownum - 1)}+Sum({item.Debit.ColumnCode(rownum)}:{item.Credit.ColumnCode(rownum)})");
                    //Common.UpdateCellDecimal(_SheetCurrentAccount, rownum, item.YearlyBalence);
                    Common.AddFormulaDecimal(_SheetCurrentAccount, rownum, item.CalculatedYearlyBalence.ColumnNumber,
                        $"={item.CalculatedYearlyBalence.ColumnCode(rownum - 1 + yearlyBalenceOffset)}+Sum({item.Debit.ColumnCode(rownum)}:{item.Credit.ColumnCode(rownum)})");
                    Common.UpdateCellString(_SheetCurrentAccount, rownum, item.Category, CategoryMissing);
                    Common.UpdateHyperLink(_SheetCurrentAccount, rownum, item.Notes, item.NotesHyperLink, stylenameHyperlink, Package.File.DirectoryName);

                    //Create conditional formating
                    if (item.Category != null)
                    {
                        int categoryColumn = item.Category.ColumnNumber;
                        ExcelAddress categoryAddress = new ExcelAddress(rownum,
                                                                        categoryColumn,
                                                                        rownum,
                                                                        categoryColumn);
                        //Wrap category text
                        _SheetCurrentAccount.Cells[categoryAddress.Address].Style.WrapText = true;
                    }
                }
                if (item.IsMonthlySummary && rownum > rowMonthStart)
                {
                    Common.AddSumFormula(_SheetCurrentAccount, rownum, colDebit, rowMonthStart, colDebit, rownum - 1, colDebit, true);
                    Common.AddSumFormula(_SheetCurrentAccount, rownum, colCredit, rowMonthStart, colCredit, rownum - 1, colCredit, true);
                    rowMonthStart = rownum + 1;

                }
                else if (item.IsDivider)
                {
                    rowMonthStart = rownum + 1;
                    yearlyBalenceOffset = -2;
                }
                else
                { yearlyBalenceOffset = 0; }
            }
        }

        private void UpsertAnnualSummary()
        {

            string stylenameHyperlink = "HyperLink";
            SummaryExpense FirstExpense = null;
            if (Expenses.Any())
            {
                FirstExpense = Expenses.First();
            }
            SummaryInvoice FirstInvoice = Invoices.First();
            Dictionary<string, ColumnHeader> chExpences = Common.GetColumnHeaders(_SheetAnnualSummary, 1, eDescriptionKeys.AnnualSummary.Expenses, 2);
            Dictionary<string, ColumnHeader> chInvoices = Common.GetColumnHeaders(_SheetAnnualSummary, 1, eDescriptionKeys.AnnualSummary.Invoices);
            Dictionary<string, ColumnHeader> chSummary = new Dictionary<string, ColumnHeader>();
            int LastExpenseColumnNumber = Common.GetLastColumnNumber(chExpences);
            int LastInvoiceColumnNumber = Common.GetLastColumnNumber(FirstInvoice);

            int rownum = 4;
            Common.DeleteRows(_SheetAnnualSummary, 4);
            #region Add Summary


            var caRows = CurrentAccountRows.Where(w => !w.IsDivider &&
                                                       !w.IsMonthlySummary &&
                                                       !w.IsStartingBalence &&
                                                       !w.IsInvoicePaid &&
                                                       !w.IsDontMap).ToList();

            caRows.Sort(delegate (CurrentAccount x, CurrentAccount y)
            {
                if (x.Description == null && y.Description == null) return 0;
                else if (x.Description == null) return -1;
                else if (y.Description == null) return 1;
                else return x.Description.Value.CompareTo(y.Description.Value);
            });

            #region Build summary Headers lookup
            int summaryColumnHeaderNumber = 0;
            List<string> SummaryHeaders = "Date,Payee,Reconsiliation,Total Spent,VAT,Dividends,Salary,Expenses,PAYE".Split(',').ToList();
            List<string> SummaryHeaders2 = new List<string>();
            foreach (CreditCard item in CreditCardRows.Where(w => !SummaryHeaders.Any(a => a.Equals(w.Category.Value, StringComparison.CurrentCultureIgnoreCase))))
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
                    item.Description.Value.Equals(Resources.Category_CommercialCard, StringComparison.CurrentCultureIgnoreCase))
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

            #endregion
            //add summary headers to sheet
            Common.SetHeaders(_SheetAnnualSummary, rownum, chSummary);

            #region add summary

            int summaryColumnsCount = chSummary.Count();
            foreach (CurrentAccount currentAccount in caRows.Where(w => !w.IsDivider &&
                                                                        !w.IsMonthlySummary &&
                                                                        !w.IsStartingBalence &&
                                                                        !w.IsInvoicePaid &&
                                                                        !w.IsDontMap &&
                                                                        w.Debit.Value != 0))
            {

                rownum++;

                Common.UpdateCellDate(_SheetAnnualSummary, rownum, new ColumnDateTime() { ColumnNumber = 1, Value = currentAccount.Date.Value });
                Common.UpdateCellString(_SheetAnnualSummary, rownum, new ColumnString() { ColumnNumber = 2, Value = currentAccount.Description.Value });
                if (!currentAccount.IsCreditCard)
                {
                    int colnum = chSummary.Single(w => w.Key.Equals(currentAccount.Category.Value, StringComparison.CurrentCultureIgnoreCase)).Value.ColumnNumber;
                    Common.UpdateCellDecimal(_SheetAnnualSummary, rownum, new ColumnDecimal() { ColumnNumber = colnum, Value = currentAccount.Debit.Value });
                }
                else if (currentAccount.CreditCardTransactionSummary != null)
                {
                    foreach (TransactionSummary ts in currentAccount.CreditCardTransactionSummary)
                    {
                        if (ts.Description == null || string.IsNullOrEmpty(ts.Description))
                        {
                            ts.Description = UnknonwnCategory;
                        }

                        int colnum = chSummary.Single(w => w.Key.Equals(ts.Description, StringComparison.CurrentCultureIgnoreCase)).Value.ColumnNumber;
                        Common.UpdateCellDecimal(_SheetAnnualSummary, rownum, new ColumnDecimal() { ColumnNumber = colnum, Value = ts.Value });
                    }
                }

                Common.AddFormulaDecimal(_SheetAnnualSummary, rownum, 3, $"ABS(D{rownum})-ABS({currentAccount.Debit.Value})");
                Common.AddSumFormula(_SheetAnnualSummary, rownum, 4, rownum, 5, rownum, summaryColumnsCount);

            }
            #endregion
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
                                        ColumnNumber = chExpences[Resources.ColumnHeader_VAT].ColumnNumber,
                                        Value = Resources.IfApplicable
                                    },
                                    "",
                                    false);


            Common.SetRowColour(_SheetAnnualSummary, rownum, LastExpenseColumnNumber, Common.Colours.HeaderColour, true);
            //////
            rownum++;
            int firstExpenseRow = rownum;
            Common.SetHeaders(_SheetAnnualSummary, rownum, chExpences);

            foreach (SummaryExpense expense in Expenses)
            {
                rownum++;
                Common.UpdateCellDate(_SheetAnnualSummary, rownum, expense.Date);
                Common.UpdateCellString(_SheetAnnualSummary, rownum, expense.Description);
                Common.UpdateCellDecimal(_SheetAnnualSummary, rownum, expense.Value);
                Common.UpdateCellDecimal(_SheetAnnualSummary, rownum, expense.VAT);
                var SumAddress = new ExcelAddress(rownum, 3, rownum, 3);
                if (expense.IsExpenseRefund)
                {
                    _SheetAnnualSummary.Cells[SumAddress.Address].Value = 0;
                }
                else
                {
                    var SumRange = new ExcelAddress(rownum, 4, rownum, chExpences.Count());
                    _SheetAnnualSummary.Cells[SumAddress.Address].Formula = $"SUM({SumRange.Address})";
                }
            }
            rownum += 2;

            //Expenses Total row
            _SheetAnnualSummary.Cells[rownum, chExpences[Resources.ColumnHeader_Description].ColumnNumber].Value = eDescriptionKeys.Totals;
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
            //Set the Invoice header
            Common.SetHeaders(_SheetAnnualSummary, rownum, chInvoices);


            //set the expense data
            int FirstInvoiceRow = rownum + 1;
            foreach (SummaryInvoice invoice in Invoices)
            {
                if (string.IsNullOrEmpty(invoice.Customer.Value)) continue;
                rownum++;
                Common.UpdateCellString(_SheetAnnualSummary, rownum, invoice.Customer);
                Common.UpdateHyperLink(_SheetAnnualSummary, rownum, invoice.InvoiceName, invoice.InvoiceNameHyperLink, stylenameHyperlink, Package.File.DirectoryName);
                Common.UpdateCellDate(_SheetAnnualSummary, rownum, invoice.InvoiceDate);
                Common.UpdateCellDate(_SheetAnnualSummary, rownum, invoice.InvoicePaid);
                Common.UpdateCellInt(_SheetAnnualSummary, rownum, invoice.DaysToPay);
                Common.UpdateCellInt(_SheetAnnualSummary, rownum, invoice.HoursInvoiced);
                Common.UpdateCellDecimal(_SheetAnnualSummary, rownum, invoice.DaysInvoiced);
                Common.UpdateCellDecimal(_SheetAnnualSummary, rownum, invoice.InvoiceAmount);
                Common.UpdateCellDecimal(_SheetAnnualSummary, rownum, invoice.TotalPaid);
                Common.UpdateCellDecimal(_SheetAnnualSummary, rownum, invoice.DayRate);
            }
            rownum += 2;
            ////Invoice Total Row
            _SheetAnnualSummary.Cells[rownum, chInvoices[Resources.ColumnHeader_Customer].ColumnNumber].Value = eDescriptionKeys.Totals;
            Common.SetTotal(_SheetAnnualSummary, rownum, FirstInvoiceRow, chInvoices[SalesInvoicesColumnheaderText.HoursInvoiced].ColumnNumber);
            Common.SetTotal(_SheetAnnualSummary, rownum, FirstInvoiceRow, chInvoices[Resources.ColumnHeader_DaysInvoiced].ColumnNumber);
            Common.SetTotal(_SheetAnnualSummary, rownum, FirstInvoiceRow, chInvoices[Resources.ColumnHeader_InvoiceAmount].ColumnNumber);
            Common.SetTotal(_SheetAnnualSummary, rownum, FirstInvoiceRow, chInvoices[Resources.ColumnHeader_TotalPaid].ColumnNumber);
            Common.SetRowColour(_SheetAnnualSummary, rownum, LastInvoiceColumnNumber, Common.Colours.TotalsColour, true);
            #endregion

            #region Add Travel Subsitance Summary
            int colText = chInvoices[SalesInvoicesColumnheaderText.HoursInvoiced].ColumnNumber;
            int colValue = chInvoices[SalesInvoicesColumnheaderText.DaysInvoiced].ColumnNumber;
            string colLetterValue = chInvoices[SalesInvoicesColumnheaderText.DaysInvoiced].GetColumnLetter();
            int availableDays = 252;
            int subsistanceDays = caRows.Where(w => w.Category.Value.Equals(Resources.Summary_Subsistence)).GroupBy(g => g.Date).ToArray().Count();

            rownum += 2;
            Common.UpdateCellString(_SheetAnnualSummary, rownum, colText, Resources.Summary_InvoicedDays   );
            Common.AddFormulaDecimal(_SheetAnnualSummary, rownum, colValue, $"={colLetterValue}{rownum - 2}");
            int invoicedDaysRow = rownum;
            rownum++;
            Common.UpdateCellString(_SheetAnnualSummary, rownum, colText, Resources.Summary_AvailableDays);
            Common.UpdateCellInt(_SheetAnnualSummary, rownum, colValue, availableDays, false, 0);
            rownum++;
            Common.UpdateCellString(_SheetAnnualSummary, rownum, colText, Resources.Summary_PercentageWorked);
            Common.AddFormulaPercentage(_SheetAnnualSummary, rownum, colValue
                                                    , rownum - 2, colLetterValue
                                                    , rownum - 1, colLetterValue);
            rownum += 2;
            Common.UpdateCellString(_SheetAnnualSummary, rownum, colText, Resources.Summary_DaysDriven);
            Common.AddFormula(_SheetAnnualSummary, rownum, colValue, $"='{Resources.WorkSheetLabel_CarMilage}'!K4", 0);
            rownum++;
            Common.UpdateCellString(_SheetAnnualSummary, rownum, colText, Resources.Summary_PercentageDriven);
            Common.AddFormulaPercentage(_SheetAnnualSummary, rownum, colValue
                                                    , rownum - 1, colLetterValue
                                                    , invoicedDaysRow, colLetterValue);

            rownum++;
            Common.UpdateCellString(_SheetAnnualSummary, rownum, colText, Resources.Summary_DaysTrain);
            Common.AddFormula(_SheetAnnualSummary, rownum, colValue, $"=RoundUp({colLetterValue}{invoicedDaysRow} - {colLetterValue}{rownum - 2},0)", 0);
            rownum++;
            Common.UpdateCellString(_SheetAnnualSummary, rownum, colText, Resources.Summary_PercentageTrain );
            Common.AddFormulaPercentage(_SheetAnnualSummary, rownum, colValue
                                                    , rownum - 1, colLetterValue
                                                    , invoicedDaysRow, colLetterValue);
            rownum++;
            Common.UpdateCellString(_SheetAnnualSummary, rownum, colText, Resources.Summary_DaysTraveled);
            Common.AddFormula(_SheetAnnualSummary, rownum, colValue, $"={colLetterValue}{rownum - 2} + {colLetterValue}{rownum - 4}", 0);
            rownum++;
            Common.UpdateCellString(_SheetAnnualSummary, rownum, colText, Resources.Summary_PercentageTraveled);
            Common.AddFormulaPercentage(_SheetAnnualSummary, rownum, colValue
                                                    , rownum - 1, colLetterValue
                                                    , invoicedDaysRow, colLetterValue);
            rownum++;
            Common.UpdateCellString(_SheetAnnualSummary, rownum, colText, Resources.Summary_DaysSubsistanceCard);
            Common.UpdateCellInt(_SheetAnnualSummary, rownum, colValue, subsistanceDays, false, 0);
            rownum++;
            Common.UpdateCellString(_SheetAnnualSummary, rownum, colText, Resources.Summary_DaysSubsistanceCash);
            Common.AddFormula(_SheetAnnualSummary, rownum, colValue, $"=RoundUp({colLetterValue}{invoicedDaysRow} - {colLetterValue}{rownum - 1},0)", 0);

            #endregion


        }

        public static class SalesInvoicesColumnheaderText
        {
            public static string Customer { get { return Resources.ColumnHeader_Customer; } }
            public static string Invoice { get { return Resources.ColumnHeader_Invoice; } }
            public static string InvoiceDate { get { return Resources.ColumnHeader_InvoiceDate; } }
            public static string DateFundsRecieved { get { return Resources.ColumnHeader_DateFundsRecieved; } }
            public static string DaysToPay { get { return Resources.ColumnHeader_DaysToPay; } }
            public static string HoursInvoiced { get { return Resources.ColumnHeader_HoursInvoiced; } }
            public static string DaysInvoiced { get { return Resources.ColumnHeader_DaysInvoiced; } }
            public static string TotalPaid { get { return Resources.ColumnHeader_TotalPaid; } }
            public static string DayRate { get { return Resources.ColumnHeader_DayRate; } }
        }
    }
}
