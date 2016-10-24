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

namespace Lava3.Core
{
    public class ExcelFile : IDisposable
    {

        public static class WorkSheetLabels
        {
            public const string CreditCard = "CreditCard";
            public const string CategoryLookup = "Category LookUp";
            public const string CurrentAccount = "HSBC";
            public const string Summary = "Annual Summary";
        }


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
        public ExcelPackage Package { get; set; }
        internal ExcelWorksheet _SheetCategories;
        internal ExcelWorksheet _SheetCreditCard;
        internal ExcelWorksheet _SheetCurrentAccount;
        public IEnumerable<Category> CategoryRows { get; set; }
        public IEnumerable<CreditCard> CreditCardRows { get; set; }
        public IList<CurrentAccount> CurrentAccountRows { get; set; }
        public Dictionary<string, dynamic> CategoryColumns { get; set; }

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
                _SheetCategories = Package.Workbook.Worksheets[WorkSheetLabels.CategoryLookup];
            }
        }

        public void ProcessCurrentAccount()
        {
            _SheetCurrentAccount = Package.Workbook.Worksheets[WorkSheetLabels.CurrentAccount];
            var columnHeaders = Common.GetColumnHeaders(_SheetCurrentAccount, 2);
            var rows = new List<CurrentAccount>();
            int rownum = 3;
            CurrentAccountRows = new List<CurrentAccount>();
            while (rownum <= _SheetCurrentAccount.Dimension.Rows)
            {
                CurrentAccount row = new CurrentAccount(_SheetCurrentAccount, columnHeaders, rownum, CategoryRows, CreditCardRows);

                rownum++;
                CurrentAccountRows.Add(row);
            }
            //Remove boundries and monthly totals
            for (int i = CurrentAccountRows.Count - 1; i >= 0; i--)
            {
                if (CurrentAccountRows[i].IsDivider || CurrentAccountRows[i].IsMonthlySummary)
                {
                    CurrentAccountRows.RemoveAt(i);
                }
            }

            //Grab the starting balance
            // sort by transaction date
            CurrentAccountRows = CurrentAccountRows.OrderBy(o => o.Date)
                                                   .ToList();
            //Set the monthly and annual running totals.
            int currentMonth = -1;
            int previousMonth = ((DateTime)CurrentAccountRows[1].Date).Month;
            Decimal? MonthlyTotal = 0m;
            for (int i = 1; i < CurrentAccountRows.Count; i++)
            {
                var previous = CurrentAccountRows[i - 1];
                var current = CurrentAccountRows[i];
                previousMonth = ((DateTime)previous.Date).Month;
                currentMonth = ((DateTime)current.Date).Month;
                if (currentMonth != previousMonth)
                {
                    MonthlyTotal = 0m;
                }
                decimal? transactionBalence = current.Credit.Value - current.Debit.Value;
                MonthlyTotal += transactionBalence;


                CurrentAccountRows[i].YearlyBalence.Value = previous.YearlyBalence.Value
                                                            + transactionBalence;
                CurrentAccountRows[i].MonthlyBalence.Value = MonthlyTotal;

            }
        }
        /// <summary>
        /// Load the credit card into memory
        /// </summary>
        public void ProcessCreditCard()
        {
            ProcessCategory();

            _SheetCreditCard = Package.Workbook.Worksheets[WorkSheetLabels.CreditCard];

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

        public void Save()
        {
            UpsertCatergory();
            UpsertCreditCard();
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
        /// <summary>
        /// Load the category into Memory
        /// </summary>
        public void ProcessCategory()
        {
            CategoryColumns = Common.GetColumnHeaders(_SheetCategories, 1);

            List<Category> accountingCategories = new List<Category>();
            int rownum = 2;
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

                        row.NotesHyperLink = _SheetCategories.Cells[rownum, CategoryColumns["Notes"].ColumnNumber].Hyperlink.OriginalUri;
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
            Common.DeleteRows(_SheetCategories);

            string stylenameHyperlink = "HyperLink";
            CreateStyleHyperLink(_SheetCategories, stylenameHyperlink);

            int rownum = 1;
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

            Common.DeleteRows(_SheetCreditCard);
            //Create styles
            string stylenameHyperlink = "HyperLink";
            CreateStyleHyperLink(_SheetCreditCard, stylenameHyperlink);

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
    }
}
