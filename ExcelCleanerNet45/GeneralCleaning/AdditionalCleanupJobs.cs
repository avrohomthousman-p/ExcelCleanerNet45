﻿using OfficeOpenXml;
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
        /// Sets all columns in the worksheet to the specified width ONLY if they are already
        /// larger than that
        /// </summary>
        /// <param name="worksheet">the worksheet being cleaned</param>
        /// <param name="minWith">the width all columns should not be larger than</param>
        internal static void ApplyColumnMaxWidth(ExcelWorksheet worksheet, double maxWith)
        {
            for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
            {
                var column = worksheet.Column(col);

                if (column.Width > maxWith)
                {
                    column.Width = maxWith;
                }
            }
        }



        /// <summary>
        /// Corrects the alignment issue that is sometimes found in the BalanceSheetDrillTrough report
        /// </summary>
        /// <param name="worksheet">the worksheet in need of realignment</param>
        internal static void RealignDataColumn(ExcelWorksheet worksheet)
        {
            //move the iterator to the top of the column
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
            ExcelRange cell = GetCellWithText(worksheet, headerText);
            if (cell != null)
            {
                cell.Style.HorizontalAlignment = desiredAlignment;
            }
        }




        /// <summary>
        /// Gets the first cell in the worksheet that has the desired text in it, or null if no such text is found
        /// </summary>
        /// <param name="worksheet">the worksheet being checked</param>
        /// <param name="targetText">the text the desired cell has</param>
        /// <returns>the row and column of the first cell containing the desired text, or null if no cell is found</returns>
        internal static ExcelRange GetCellWithText(ExcelWorksheet worksheet, string targetText)
        {
            ExcelIterator iter = new ExcelIterator(worksheet);
            return iter.GetFirstMatchingCell(cell => cell.Text.Trim() == targetText);
        }




        /// <summary>
        /// Finds the column with the specified header and turns on WrapText in all cells in that column. 
        /// </summary>
        /// <param name="worksheet">the worksheet being cleaned</param>
        /// <param name="columnHeader">the literal text found at the header of that column</param>
        internal static void SetColumnToWrapText(ExcelWorksheet worksheet, string columnHeader)
        {
            var topOfColumn = GetCellWithText(worksheet, columnHeader);
            if (topOfColumn == null) //if no such column exists
            {
                return;
            }


            SetColumnToWrapText(worksheet, topOfColumn.Start.Row + 1, topOfColumn.Start.Column);
        }




        /// <summary>
        /// Finds the column with the specified header and ensures its width is no less than the desired width. 
        /// </summary>
        /// <param name="worksheet">the worksheet being cleaned</param>
        /// <param name="columnHeader">the literal text found at the header of that column</param>
        /// <param name="desiredWidth">the minimum width to be applied to the column</param>
        internal static void SetColumnToMinimumWidth(ExcelWorksheet worksheet, string columnHeader, double desiredWidth)
        {
            var topOfColumn = GetCellWithText(worksheet, columnHeader);
            if (topOfColumn == null) //if no such column exists
            {
                return;
            }


            SetColumnToMinimumWidth(worksheet, topOfColumn.Start.Column, desiredWidth);
        }




        /// <summary>
        /// Ensures the specified column has a width that is no less than the desired width. 
        /// </summary>
        /// <param name="worksheet">the worksheet being cleaned</param>
        /// <param name="column">the column number of the column to be resized</param>
        /// <param name="desiredWidth">the minimum width to be applied to the column</param>
        internal static void SetColumnToMinimumWidth(ExcelWorksheet worksheet, int column, double desiredWidth)
        {
            if(worksheet.Column(column).Width < desiredWidth)
            {
                worksheet.Column(column).Width = desiredWidth;
            }
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




        /// <summary>
        /// Sets the size of the specified column to the size of itself plus the next two columns, 
        /// and then deletes those next two columns.
        /// 
        /// 
        /// This is useful for the VendorInvoiceReport which has 2 extra columns that need to be deleted
        /// </summary>
        /// <param name="worksheet">the worksheet being cleaned</param>
        /// <param name="column">the column that will be resized</param>
        internal static void ResizeColumnAndDeleteTheNext2(ExcelWorksheet worksheet, int column)
        {
            double totalSize = worksheet.Column(column).Width;

            

            if (SafeToDeleteColumn(worksheet, column + 1))
            {
                totalSize += worksheet.Column(column + 1).Width;
                worksheet.DeleteColumn(column + 1);
            }
            if (SafeToDeleteColumn(worksheet, column + 1))
            {
                totalSize += worksheet.Column(column + 1).Width;
                worksheet.DeleteColumn(column + 1);
            }


            worksheet.Column(column).Width = totalSize;
        }





        /// <summary>
        /// Checks if a column is safe to delete becuase it is empty other than possibly having major headers in it.
        /// </summary>
        /// <param name="worksheet">the worksheet where the column can be found</param>
        /// <param name="col">the column being checked</param>
        /// <returns>true if it is safe to delete the column and false if deleting it would result in data loss</returns>
        internal static bool SafeToDeleteColumn(ExcelWorksheet worksheet, int col)
        {
            for (int row = 1; row <= worksheet.Dimension.End.Row; row++)
            {
                if (!FormulaManager.IsEmptyCell(worksheet.Cells[row, col]))
                {
                    return false;
                }
            }


            return true;
        }




        /// <summary>
        /// Deletes all empty cells in the DistributionsReport
        /// </summary>
        /// <param name="worksheet">the worksheet being cleaned</param>
        /// <param name="topHeader">the header signaling the first row that needs cells deleted</param>
        /// <param name="bottomHeader">the header signaling the last row the has cells that need to be deleted</param>
        internal static void DeleteEmptyCellsForDistributionsReport(ExcelWorksheet worksheet, string topHeader, string bottomHeader)
        {
            //Find top row
            var topCell = GetCellWithText(worksheet, topHeader);
            if(topCell == null)
            {
                Console.WriteLine("no top cell found");
                return;
            }

            int topRow = topCell.Start.Row;


            //Find bottom row
            ExcelIterator iter = new ExcelIterator(worksheet);
            Tuple<int, int> bottomCell = iter
                            .FindAllMatchingCoordinates(c => FormulaManager.TextMatches(c.Text, bottomHeader))
                            .LastOrDefault();


            if(bottomCell == null)
            {
                Console.WriteLine("no bottom found");
                return;
            }


            if(!AllCellsAreEmpty(worksheet, topRow, bottomCell.Item1, bottomCell.Item2))
            {
                Console.WriteLine("thats not empty");
                return;
            }


            //delete the cells
            ExcelRange cell;
            for(int i = topRow; i < bottomCell.Item1; i++)
            {
                cell = worksheet.Cells[i, bottomCell.Item2];
                cell.Delete(eShiftTypeDelete.Left);
            }


            //also delete cell to the right and left of the bottom cell to better align the summary
            cell = worksheet.Cells[bottomCell.Item1, bottomCell.Item2 + 1];
            cell.Delete(eShiftTypeDelete.Left);
            cell = worksheet.Cells[bottomCell.Item1, bottomCell.Item2 - 1];
            cell.Delete(eShiftTypeDelete.Left);
        }



        /// <summary>
        /// Checks if all cells in the specified column and between the specified top and bottom
        /// are empty
        /// </summary>
        /// <param name="worksheet">the worksheet being cleaned</param>
        /// <param name="top">the top row to be checked (inclusive)</param>
        /// <param name="bottom">the bottom row to be checked (exclusive)</param>
        /// <param name="col">the column to be checked</param>
        /// <returns></returns>
        private static bool AllCellsAreEmpty(ExcelWorksheet worksheet, int top, int bottom, int col)
        {
            ExcelRange cell;

            bool allEmpty = true;
            for(int i = top; i < bottom; i++)
            {
                cell = worksheet.Cells[i, col];

                if (!FormulaManager.IsEmptyCell(cell))
                {
                    allEmpty = false;
                    break;
                }
            }

            return allEmpty;
        }



        /// <summary>
        /// Sets all columns in the worksheet to wrap their text to the next line.
        /// </summary>
        /// <param name="worksheet">the worksheet being cleaned</param>
        /// <param name="numCellsInFirstRow">the number of cells to expect in the first data row</param>
        internal static void SetAllColumnsToWrapText(ExcelWorksheet worksheet)
        {
            int firstRow = FindFirstRowWithMultipleEntries(worksheet, 4);
            if(firstRow == -1)
            {
                return;
            }

            SetAllColumnsToWrapText(worksheet, firstRow);
        }



        /// <summary>
        /// Sets all the columns in the worksheet to wrp the text to the next line
        /// </summary>
        /// <param name="worksheet">the worksheet being cleaned</param>
        /// <param name="firstDataRow">the topmost row that should be set to wrap its text</param>
        internal static void SetAllColumnsToWrapText(ExcelWorksheet worksheet, int firstDataRow)
        {
            for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
            {
                SetColumnToWrapText(worksheet, firstDataRow, col);
            }
        }




        /// <summary>
        /// Finds the first row in the worksheet that has a number of non empty cells greater than or equal to 
        /// the passed in number of required entries.
        /// </summary>
        /// <param name="worksheet">the worksheet being cleaned</param>
        /// <param name="numRequiredEntries">the number of non empty cells a row must have</param>
        /// <returns>the row number of the first row with sufficent entries, or -1 if no such cell is found</returns>
        internal static int FindFirstRowWithMultipleEntries(ExcelWorksheet worksheet, int numRequiredEntries)
        {
            ExcelIterator iter = new ExcelIterator(worksheet);
            for(int i = 1; i <= worksheet.Dimension.End.Row; i++)
            {
                iter.SetCurrentLocation(i, 1);

                int nonEmptyCells = iter.GetCells(ExcelIterator.SHIFT_RIGHT)
                                        .Count(cell => !FormulaManager.IsEmptyCell(cell));


                if(nonEmptyCells >= numRequiredEntries)
                {
                    return i;
                }
            }


            return -1;
        }




        /// <summary>
        /// Fixes the misaligned summary cells at the bottom of the InvoiceList report
        /// </summary>
        /// <param name="worksheet">the worksheet being cleaned</param>
        internal static void FixSummariesOfInvoiceList(ExcelWorksheet worksheet)
        {
            //Deletes some empty cells to move the data cells to where they should be

            worksheet.DeleteRow(worksheet.Dimension.End.Row);
            ExcelRange cell = worksheet.Cells[worksheet.Dimension.End.Row - 1, worksheet.Dimension.End.Column - 3, 
                                                worksheet.Dimension.End.Row - 1, worksheet.Dimension.End.Column - 1];

            cell.Delete(eShiftTypeDelete.Up);


            worksheet.DeleteRow(worksheet.Dimension.End.Row);
            cell = worksheet.Cells[worksheet.Dimension.End.Row, worksheet.Dimension.End.Column - 1];
            cell.Delete(eShiftTypeDelete.Left);
        }



        /// <summary>
        /// Widens most of the columns in the ReportTenantSummary
        /// </summary>
        /// <param name="worksheet">the worksheet being cleaned</param>
        internal static void SetColumnWidthsForReportTenantSummary(ExcelWorksheet worksheet)
        {
            const double DEFAULT_WIDTH = 12;
            ExcelColumn col;

            for (int i = 4; i <= worksheet.Dimension.End.Column; i++)
            {
                col = worksheet.Column(i);
                if (col.Width < DEFAULT_WIDTH)
                {
                    col.Width = DEFAULT_WIDTH;
                }
            }
        }




        /// <summary>
        /// Ensures that all major headers in the first row are left aligned. A major header here, refers
        /// to any text that appears before the first row in the worksheet with at least 3 non empty cells in it
        /// </summary>
        /// <param name="worksheet">the worksheet that needs to be cleaned</param>
        internal static void LeftAlignAllHeadersInFirstRow(ExcelWorksheet worksheet)
        {
            int tableStart = FindFirstRowWithMultipleEntries(worksheet, 3);

            ExcelRange cell;
            for(int i = 1; i < tableStart; i++)
            {
                cell = worksheet.Cells[i, 1];
                cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            }
        }
    }
}
