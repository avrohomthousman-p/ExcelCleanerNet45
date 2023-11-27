using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Threading.Tasks;

namespace ExcelCleanerNet45
{
    /// <summary>
    /// Implementation of IFormulaGenerator that looks for the specifed header followed by the other specified header
    /// (e.g. "Income" and "Total Income"), and treats the rows between those headers as a "formula range" that 
    /// gets its own formula. Each header pair should be included in the string array passed to InsertFormulas
    /// in this format:  [text of start header]=[text of end header].
    /// 
    /// Note: for data cells to be considered part of the formula range they must be between the start header
    /// and the end header IN TERMS OF ROW NUMBER ONLY. The column of the start and end header, as well as all columns 
    /// to the right of that, will recieve summaries (assuming they contain data cells).
    /// </summary>
    internal class RowSegmentFormulaGenerator : IFormulaGenerator
    {

        protected IsDataCell isDataCell = new IsDataCell(FormulaManager.IsDollarValue);


        private bool trimRange = true;
        /// <summary>
        /// controls whether or not the system skips the whitespace immideatly after the top header of each formula range
        /// e.g. if we are summing up all cells from "Income" to "Total Income", this boolean decides if the formula 
        /// should include the empty cells just below the "Income" row
        /// </summary>
        public bool trimFormulaRange 
        { 
            get { return trimRange; } 
            set { trimRange = value; } 
        }




        /// <inheritdoc/>
        public void SetDataCellDefenition(IsDataCell altImplementation)
        {
            this.isDataCell = altImplementation;
        }



        /// <inheritdoc/>
        public virtual void InsertFormulas(ExcelWorksheet worksheet, string[] headers)
        {
            string startHeader, endHeader;

            foreach (string header in headers)              //for each header in the report that needs a formula 
            {

                //Ensure that the header was intended for this class and not the DistantRowsFormulaGenerator class
                if (FormulaManager.IsNonContiguousFormulaRange(header))
                {
                    continue;
                }



                int seperator = header.IndexOf('=');

                startHeader = header.Substring(0, seperator);
                endHeader = header.Substring(seperator + 1);

                var ranges = GetRowRangeForFormula(worksheet, startHeader, endHeader);

                foreach (var item in ranges)                // for each instance of that header
                {
                    FillInFormulas(worksheet, item.Item1, item.Item2, item.Item3);
                }
            }

        }




        /// <summary>
        /// Gets the row numbers of the first and last rows that should be included in the formula
        /// </summary>
        /// <param name="worksheet">the worksheet currently being given formulas</param>
        /// <param name="startHeader">the text to look for to signal the start row</param>
        /// <param name="endHeader">The text to look for to signal the end row</param>
        /// <returns>a tuple containing the start-row, end-row, and column of the formula range</returns>
        protected virtual IEnumerable<Tuple<int, int, int>> GetRowRangeForFormula(ExcelWorksheet worksheet, string startHeader, string endHeader)
        {
            ExcelIterator iter = new ExcelIterator(worksheet);

            Predicate<ExcelRange> cellMatchesStartingHeader = (cell => FormulaManager.TextMatches(cell.Text, startHeader));

            var cellsInWorksheet = iter.FindAllMatchingCoordinates( cellMatchesStartingHeader );

            foreach (Tuple<int, int> cell in cellsInWorksheet)
            {
                //search for end of sequence
                int end = FindEndOfFormulaRange(worksheet, cell.Item1, cell.Item2, endHeader);

                if (end > 0)
                {
                    yield return new Tuple<int, int, int>(cell.Item1, end, cell.Item2); //return startRow, endRow, column


                    //if we are not on the last row (index out of bounds check)
                    if (end < worksheet.Dimension.End.Row)
                    {
                        iter.SetCurrentLocation(end /* + 1*/, 1); //skip to the the next row (we dont expect 2 headers on one row)
                    }
                    
                }
            }
        }




        /// <summary>
        /// Given the Cell coordinates of the starting cell in a formula range, finds the ending cell for 
        /// that range.
        /// </summary>
        /// <param name="worksheet">the worksheet currently being given formulas</param>
        /// <param name="row">the row number of the starting cell in the formula range</param>
        /// <param name="col">the column number of the starting cell in the formula range</param>
        /// <param name="targetText">the text to look for that signals the end cell of the formula range</param>
        /// <returns>the row number of the last cell in the formula range</returns>
        protected static int FindEndOfFormulaRange(ExcelWorksheet worksheet, int row, int col, string targetText)
        {
            ExcelIterator iter = new ExcelIterator(worksheet, row + 1, col);

            Predicate<ExcelRange> matchesEndHeader = (cell => FormulaManager.TextMatches(cell.Text, targetText));

            Tuple<int, int> endCell = iter.GetCellCoordinates(ExcelIterator.SHIFT_DOWN, stopIf:matchesEndHeader).Last();
            
            return Math.Min(endCell.Item1 + 1, worksheet.Dimension.End.Row);
        }




        /// <summary>
        /// Inserts the formulas in each cell in the formula range that requires it.
        /// </summary>
        /// <param name="worksheet">the worksheet currently being given formulas</param>
        /// <param name="startRow">the first row of the formula range (containing the header)</param>
        /// <param name="endRow">the last row of the formula range (containing the total)</param>
        /// <param name="col">the column of the header and total for the formula range</param>
        protected virtual void FillInFormulas(ExcelWorksheet worksheet, int startRow, int endRow, int col)
        {

            ExcelRange cell;



            //Often there are multiple columns that require a formula, so we need to iterate
            //and apply the formulas in many columns
            for (col++; col <= worksheet.Dimension.End.Column; col++)
            {
                cell = worksheet.Cells[endRow, col];

                if (this.isDataCell(cell))
                {
                    if (trimRange)
                    {     
                        startRow += CountEmptyCellsOnTop(worksheet, startRow, endRow, col); //Skip the whitespace on top
                    }
                    string formula = FormulaManager.GenerateFormula(worksheet, startRow, endRow - 1, col);

                    FormulaManager.PutFormulaInCell(cell, formula);
                }
                else if (!FormulaManager.IsEmptyCell(cell))
                {
                    return;
                }
            }

        }



        /// <summary>
        /// Counts the number of empty cells between the start header(inclusive) and the actual data cells in the 
        /// formula range.
        /// </summary>
        /// <param name="worksheet">the worksheet in need of formulas</param>
        /// <param name="startRow">the row where the start header was found</param>
        /// <param name="endRow">te row where the end header was found</param>
        /// <param name="col">the column of the formula range</param>
        /// <returns>the number of empty cells at the start of the formula range</returns>
        protected static int CountEmptyCellsOnTop(ExcelWorksheet worksheet, int startRow, int endRow, int col)
        {
            int emptyCells = 0;
            ExcelRange cell;

            for(;startRow <= endRow; startRow++)
            {
                cell = worksheet.Cells[startRow, col];
                if (FormulaManager.IsEmptyCell(cell))
                {
                    emptyCells++;
                }
                else
                {
                    break;
                }
            }


            return emptyCells;
        }
    }
}
