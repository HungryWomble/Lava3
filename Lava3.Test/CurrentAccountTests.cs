using System;
using System.Collections.Generic;
using System.Diagnostics;
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
    public class CurrentAccountTests
    {

        [TestCase]
        public void ProcessCurrentAccount01()
        {

            var y = System.Diagnostics.Process.GetProcessesByName("EXCEL");
           

            Assert.Inconclusive();
            //string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Testfiles\\test.xlsx");
            //var target = new ExcelFile();

            //target.OpenPackage(path);
            //target.ProcessCurrentAccount();
            //List<CreditCard> actualCreditCardRows = target.CreditCardRows.ToList();
            //List<CurrentAccount> actualCurrentAccountRows = target.CurrentAccountRows.ToList();
            //target.SaveAndClose();
            ////Check the credit card
            //Assert.IsNull(target.Package);
            //Assert.AreEqual(10, actualCreditCardRows.Count);
            //Assert.AreEqual(actualCreditCardRows[0].Category.Value, string.Empty);
            //Assert.AreEqual(actualCreditCardRows[2].Category.Value, "Insurance");
            //Assert.AreEqual(actualCreditCardRows[4].Category.Value, "Training");
            //Assert.AreEqual(actualCreditCardRows[4].Notes.Value, "Plural sight");
            //Assert.AreEqual(actualCreditCardRows[4].NotesHyperLink.OriginalString,@"..\Reciepts\20130806_PluralSight.pdf");

            ////Check the Current account
            //Assert.AreEqual(7, actualCurrentAccountRows.Count);

            //Assert.AreEqual(76.38m, actualCurrentAccountRows[1].Balence.Value);
            //Assert.AreEqual(75.42m, actualCurrentAccountRows[2].Balence.Value);
            //Assert.AreEqual(72.32m, actualCurrentAccountRows[3].Balence.Value);
            //Assert.AreEqual(1072.32m, actualCurrentAccountRows[4].Balence.Value);
            //Assert.AreEqual(1037.33m, actualCurrentAccountRows[5].Balence.Value);
            //Assert.AreEqual(1022.83m, actualCurrentAccountRows[6].Balence.Value);

            //Assert.AreEqual(-23.62m, actualCurrentAccountRows[1].MonthlyBalence.Value);
            //Assert.AreEqual(-24.58m, actualCurrentAccountRows[2].MonthlyBalence.Value);
            //Assert.AreEqual(-3.10m, actualCurrentAccountRows[3].MonthlyBalence.Value);
            //Assert.AreEqual(996.9m, actualCurrentAccountRows[4].MonthlyBalence.Value);
            //Assert.AreEqual(961.91m, actualCurrentAccountRows[5].MonthlyBalence.Value);
            //Assert.AreEqual(947.41m, actualCurrentAccountRows[6].MonthlyBalence.Value);


            //for (int i = 1; i < 7; i++)
            //{
            //    Assert.AreEqual(actualCurrentAccountRows[i].Balence.Value, actualCurrentAccountRows[i].YearlyBalence.Value);
            //}


        }
    }
}
