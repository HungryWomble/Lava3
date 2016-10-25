using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Lava3.Core.DataTypes;
using OfficeOpenXml;
using System.Globalization;

namespace Lava3.Core
{
    public static class Common
    {
        public static Dictionary<string, dynamic> GetColumnHeaders(ExcelWorksheet sheet, int headerRowNumber)
        {
            if (sheet == null) throw new ArgumentNullException("sheet passed in is null");
            Dictionary<string, dynamic> retval = new Dictionary<string, dynamic>();
            int colnum = 1;
            while (colnum <= sheet.Dimension.Columns)
            {
                string key = Common.ReplaceNullOrEmpty(sheet.Cells[headerRowNumber, colnum].Value);
                if (!string.IsNullOrEmpty(key))
                {
                    retval.Add(key, new
                    {
                        Header = key,
                        ColumnNumber = colnum
                    });
                }
                colnum++;
            }
            return retval;

        }

        #region Comments
        /// <summary>
        /// Clear the workwheets
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns></returns>
        public static void ClearComments( ExcelWorksheet sheet)
        {
            if (sheet.Comments.Count > 0)
            {
                for (int i = sheet.Comments.Count; i > 0; i--)
                {
                    sheet.Comments.RemoveAt(i - 1);
                }
            }
        }

        /// <summary>
        /// Add a comment to the cell
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="address"></param>
        /// <param name="commentText"></param>
        public static void SetComment( ExcelWorksheet sheet, ExcelCellAddress cellAddress, string commentText, Color? fillColour = null)
        {
            var cell = sheet.Cells[cellAddress.Address];
            if (cell.Comment == null &&
               !string.IsNullOrEmpty(commentText))
            {
                cell.AddComment(commentText, "Lava");
            }
            else if (cell.Comment != null)
            {
                cell.Comment.Text = commentText;
                cell.Comment.Author = "Lava";
            }
            if (fillColour != null)
            {
                cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                cell.Style.Fill.BackgroundColor.SetColor((Color)fillColour);
            }
        }


        internal static void SetComment( ExcelWorksheet sheet, int rownum, int columnNumber, string commentText, System.Drawing.Color? fillColour = null)
        {
            ExcelCellAddress cellAddress = new ExcelCellAddress(rownum, columnNumber);
            SetComment( sheet, cellAddress, commentText, fillColour);
        }
        #endregion
        #region ReplaceNullOrEmpty
        public static string ReplaceNullOrEmpty(object o)
        {
            if (o == null)
            {
                return string.Empty;
            }
            return o.ToString().Trim();
        }
        public static DateTime? ReplaceNullOrEmptyDateTime(object o)
        {
            if (o == null)
            {
                return null;
            }
            if (o is double)
            {
                return DateTime.FromOADate((double)o);
            }
            return Convert.ToDateTime(o);
        }
        public static decimal? ReplaceNullOrEmptyDecimal(object o)
        {
            if (o == null)
            {
                return null;
            }
            return Convert.ToDecimal(o);
        }
        #endregion
        #region ReplaceIfEmpty
        public static string ReplaceIfEmpty(string original, string replacement)
        {
            if (!string.IsNullOrEmpty(original)) return original;
            return replacement;
        }

        internal static ColumnString ReplaceIfEmpty(ColumnString original, ColumnString replacement)
        {

            if (!string.IsNullOrEmpty(original.Value)) return original;
            original.Value = replacement.Value;
            original.Errors = replacement.Errors;
            return original;
        }
        #endregion

        #region Colours
        public static class Colours
        {
            public static System.Drawing.Color DuplicateColour { get { return System.Drawing.Color.LightGreen; } }
            public static System.Drawing.Color ErrorColour { get { return System.Drawing.Color.Red; } }
        }



        #endregion

        internal static void DeleteRows( ExcelWorksheet sheet, int startingRow = 1)
        {
            ClearComments( sheet);
            for (int i = sheet.Dimension.Rows; i > startingRow; i--)
            {
                sheet.DeleteRow(i);
            }
        }
        #region update cell

        public static void UpdateCellDate( ExcelWorksheet sheet, int rownumber, ColumnDateTime field)
        {
            if (field?.Value==null) return;
            sheet.Cells[rownumber, field.ColumnNumber].Value = ((DateTime)field.Value).ToOADate();
            sheet.Cells[rownumber, field.ColumnNumber].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
            if (field.Errors.Any())
            {
                foreach (string item in field.Errors)
                {
                    SetComment(sheet, rownumber, field.ColumnNumber, item, Colours.ErrorColour);
                }
            }
        }
        public static void UpdateCellString( ExcelWorksheet sheet, int rownumber, ColumnString field)
        {
            if (field==null || string.IsNullOrEmpty(field.Value)) return;
            sheet.Cells[rownumber, field.ColumnNumber].Value = field.Value.Trim().ToString();

            if (field.Errors.Any())
            {
                foreach (string item in field.Errors)
                {
                    SetComment(sheet, rownumber, field.ColumnNumber, item, Colours.ErrorColour);
                }
            }
        }
        public static void UpdateCellDecimal( ExcelWorksheet sheet, int rownumber, ColumnDecimal field)
        {
            if (field?.Value==null) return;
            sheet.Cells[rownumber, field.ColumnNumber].Value = field.Value;

            if (field.Errors.Any())
            {
                foreach (string item in field.Errors)
                {
                    SetComment(sheet, rownumber, field.ColumnNumber, item, Colours.ErrorColour);
                }
            }
        }

        internal static void UpdateHyperLink(ExcelWorksheet sheet, 
                                            int rownum, 
                                            ColumnString cell, 
                                            Uri hyperLink,
                                            string stylenameHyperlink)
        {
            if (hyperLink == null)
            {
                Common.UpdateCellString(sheet, rownum, cell);
            }
            else
            {
                ExcelRange cellRange = sheet.Cells[rownum, cell.ColumnNumber];
                cellRange.Hyperlink = hyperLink;
                cellRange.StyleName = stylenameHyperlink;
                cellRange.Value = cell.Value;
            }
        }


        #endregion
    }
}