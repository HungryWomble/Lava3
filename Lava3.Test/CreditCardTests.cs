using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Lava3.Core;
using Lava3.Core.Model;
using NUnit;
using NUnit.Framework;

namespace Lava3.Test
{
    [TestFixture]
    public class CreditCardTests
    {
        [TestCase]
        public void ProcessCreditCard01()
        {
            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Testfiles\\ProcessCreditCard01.xlsx");
            File.Copy(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Testfiles\\test.xlsx"), path, true);
            var target = new ExcelFile();

            target.OpenPackage(path);
            target.ProcessCreditCard();
            List<CreditCard> actual = target.CreditCardRows.ToList();
            foreach (var item in actual)
            {
                Assert.AreEqual(1, item.PaidDate.ColumnNumber);
                Assert.AreEqual(2, item.StatementDate.ColumnNumber);
                Assert.AreEqual(3, item.TransactionDate.ColumnNumber);
                Assert.AreEqual(4, item.TransactionDescription.ColumnNumber);
                Assert.AreEqual(5, item.TransactionAmount.ColumnNumber);
                Assert.AreEqual(6, item.Category.ColumnNumber);
                Assert.AreEqual(7, item.VatContent.ColumnNumber);
                Assert.AreEqual(8, item.Postage.ColumnNumber);
                Assert.AreEqual(9, item.Notes.ColumnNumber);
            }
            target.SaveAndClose();

            target = new ExcelFile();
            target.ShowFile(path);


            Assert.IsNull(target.Package);
            Assert.IsTrue(string.IsNullOrEmpty(actual[0].Category.Value));
            Assert.AreEqual(actual[2].Category.Value, "Insurance");
            Assert.AreEqual(actual[4].Category.Value, "Training");
            Assert.AreEqual(actual[4].Notes.Value, "Plural sight");
            Assert.AreEqual(actual[4].NotesHyperLink.OriginalString, @"Dilbert%2001.pdf");

        }
    }
}
