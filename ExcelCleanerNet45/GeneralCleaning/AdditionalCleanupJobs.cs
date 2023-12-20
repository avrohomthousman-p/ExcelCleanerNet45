using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelCleanerNet45.GeneralCleaning
{
    /// <summary>
    /// A place to store all methods that can be used as Additional cleanup jobs to be passed
    /// to a merge cleaner via the MergeCleaner.AddCleanupJob function.
    /// </summary>
    internal static class AdditionalCleanupJobs
    {

        /// <summary>
        /// Sets all columns in the worksheet to the specified width ONLY if they are not already
        /// larger than that
        /// </summary>
        /// <param name="worksheet">the worksheet being cleaned</param>
        /// <param name="minWith">the width all columns should not be smaller than</param>
        internal static void SetColumnsToMinimumWidth(ExcelWorksheet worksheet, double minWith)
        {
            for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
            {
                var column = worksheet.Column(col);

                if (column.Width < minWith)
                {
                    column.Width = minWith;
                }
            }
        }



        /// <summary>
        /// Corrects the alignment issue that is sometimes found in the BalanceSheetDrillTrough report
        /// </summary>
        /// <param name="worksheet">the worksheet in need of realignment</param>
        internal static void RealignDataColumn(ExcelWorksheet worksheet)
        {
            //find the top of the column
            ExcelIterator iter = new ExcelIterator(worksheet);
            ExcelRange temp = iter.GetFirstMatchingCell(cell => FormulaManager.IsDollarValue(cell));
            if (temp == null)
            {
                return;
            }

            
            //Later we will need to iterate starting from the top of the column
            ExcelIterator copy = new ExcelIterator(iter);


            //find the alignment used by the majority of cells in the column
            var mostCommonAlignment = iter.GetCells(ExcelIterator.SHIFT_DOWN)
                                        .GroupBy(cell => cell.Style.HorizontalAlignment)
                                        .OrderByDescending(group => group.Count())
                                        .FirstOrDefault().Key;


            //set all the data cells in the column to that alignment
            var dataCells = copy.GetCells(ExcelIterator.SHIFT_DOWN)
                                .Where(cell => FormulaManager.IsDollarValue(cell));

            foreach (ExcelRange cell in dataCells)
            {
                cell.Style.HorizontalAlignment = mostCommonAlignment;
            }
        }



        /// <summary>
        /// Searches the worksheet for a header that matches the specifeid text and gives it the specified alignment
        /// </summary>
        /// <param name="worksheet">the worksheet in need of cleaning</param>
        /// <param name="headerText">the text the header has</param>
        /// <param name="desiredAlignment">the alignment to be assigned to the header</param>
        internal static void RealignSingleHeader(ExcelWorksheet worksheet, string headerText, ExcelHorizontalAlignment desiredAlignment)
        {
            ExcelIterator iter = new ExcelIterator(worksheet);
            ExcelRange cell = iter.GetFirstMatchingCell(c => c.Text.Trim() == headerText);
            if (cell != null)
            {
                cell.Style.HorizontalAlignment = desiredAlignment;
            }
        }



        /// <summary>
        /// Finds the column with the specified header and turns on WrapText in all cells in that column. 
        /// </summary>
        /// <param name="worksheet">the worksheet being cleaned</param>
        /// <param name="columnHeader">the literal text found at the header of that column</param>
        internal static void SetColumnToWrapText(ExcelWorksheet worksheet, string columnHeader)
        {
            ExcelIterator iter = new ExcelIterator(worksheet);
            var topOfColumn = iter.GetFirstMatchingCell(cell => cell.Text.Trim() == columnHeader);
            if (topOfColumn == null) //if no such column exists
            {
                return;
            }


            SetColumnToWrapText(worksheet, topOfColumn.End.Row + 1, topOfColumn.Start.Column);
        }




        /// <summary>
        /// Sets all cells in the column below the specifed starting point (inclusive) to WrapText
        /// </summary>
        /// <param name="worksheet">the worksheet being cleaned</param>
        /// <param name="row">the topmost row in the desired column</param>
        /// <param name="col">the column number of the desired column</param>
        internal static void SetColumnToWrapText(ExcelWorksheet worksheet, int row, int col)
        {
            ExcelIterator iter = new ExcelIterator(worksheet, row, col);

            var cells = iter.GetCells(ExcelIterator.SHIFT_DOWN);

            foreach (ExcelRange cell in cells)
            {
                cell.Style.WrapText = true;
            }
        }
    }
}
