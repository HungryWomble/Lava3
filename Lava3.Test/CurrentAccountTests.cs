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


            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Testfiles\\ProcessCurrentAccount01.xlsx");
            File.Copy(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Testfiles\\test.xlsx"), path, true);
            var target = new ExcelFile();

            target.OpenPackage(path);
            target.ProcessCurrentAccount();
            List<CreditCard> actualCreditCardRows = target.CreditCardRows.ToList();
            List<CurrentAccount> actualCurrentAccountRows = target.CurrentAccountRows.ToList();
            target.SaveAndClose();
            //Check the credit card
            Assert.IsNull(target.Package);
           


            //for (int i = 1; i < 7; i++)
            //{
            //    Assert.AreEqual(actualCurrentAccountRows[i].Balence.Value, actualCurrentAccountRows[i].YearlyBalence.Value);
            //}


            target = new ExcelFile();
            target.ShowFile(path);

            Assert.Inconclusive();

        }
    }
}
