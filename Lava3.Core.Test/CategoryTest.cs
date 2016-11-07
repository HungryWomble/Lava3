using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Lava3.Core.Model;
using NUnit;
using NUnit.Framework;
using System.Text.RegularExpressions;
using Lava3.Core;
using System.IO;
using Lava3.Core.DataTypes;
using OfficeOpenXml;

namespace Lava3.Test
{
    [TestFixture]
    public class CategoryTest
    {
        [TestCase("strikes bAck", "^(strikes bAck)$", "strikes bAck", 1)]
        [TestCase("strikes bAck", "^(strikes bAck)$", "The empire strikes bAck", 0)]
        [TestCase("*strikes bAck", "(strikes bAck)$", "The empire strikes bAck", 1)]
        [TestCase("*strikes bAck", "(strikes bAck)$", "The empire strikes bAck again", 0)]
        [TestCase("strikes bAck*", "^(strikes bAck)", "strikes bAck", 1)]
        [TestCase("strikes bAck*", "^(strikes bAck)", "the empire strikes bAck", 0)]
        [TestCase("strikes bAck*", "^(strikes bAck)", "strikes bAck again", 1)]
        [TestCase("*strikes bAck*", "(strikes bAck)", "The empire strikes bAck again", 1)]
        [TestCase("*strikes bAck*", "(strikes bAck)", "strikes bAck", 1)]
        [TestCase("*strikes bAck*", "(strikes bAck)", "The empire strikes", 0)]
        public void BuildRegex01(string description, string expectedregex, string testPhase, int expectedmatches)
        {
            var target = new Category();
            target.Description = new ColumnString() { Value = description };
            Assert.AreEqual(expectedregex, target.RegEx.Value);

            var regex = new Regex(target.RegEx.Value, RegexOptions.IgnoreCase);
            var actual = regex.Match(testPhase).Length;
            if (actual > 1)
            {
                actual = 1;
            }
            Assert.AreEqual(expectedmatches, actual);

        }

        [TestCase]
        public void ProcessCategories01()
        {

            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Testfiles\\ProcessCategories01.xlsx");
            File.Copy(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Testfiles\\test.xlsx"), path, true);
            var target = new ExcelFile();

            target.OpenPackage(path);
            target.LoadCategory();
            List<Category> actual = target.CategoryRows.ToList();
            target.SaveAndClose();
            Assert.IsNull(target.Package);


            using (var Package = new ExcelPackage(new FileInfo(path)))
            {
                ExcelWorksheet sheet = Package.Workbook.Worksheets[ExcelFile.eWorkSheetLabels.CategoryLookup];
                Assert.AreEqual("Plural sight", sheet.Cells[2, 3].Text);
                Assert.IsNotNull(sheet.Cells[3,1].Comment.Text);
            }
            //target = new ExcelFile();
            //target.ShowFile(path);
        }
    }
}

