using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelCleanerNet45.FormulaGeneration.ReportSpecificGenerators
{

    internal delegate bool IsOutsideFormula(ExcelRange cell);
    internal delegate bool StopGivingFormulas(ExcelRange cell);


    /// <summary>
    /// Implementation of IFormulaGenerator that gives formulas to a column that is the sum of all data
    /// columns to its left. This is similar to the SummaryColumnGenerator, except that it adds
    /// all columns to the left, instead of just adding specific columns. Also, columns cannot be made negetive.
    /// 
    /// The header argument for this formula generator is just the header found at the top of the column that should
    /// contain all the formulas.
    /// </summary>
    internal class FullTableSummaryColumn : IFormulaGenerator
    {
        private IsDataCell dataCellDef;
        private IsOutsideFormula outsideFormula;
        private StopGivingFormulas columnEnds;



        public FullTableSummaryColumn()
        {
            dataCellDef = new IsDataCell(cell => FormulaManager.IsDollarValue(cell));
            outsideFormula = new IsOutsideFormula(cell => !FormulaManager.IsDollarValue(cell));
            columnEnds = new StopGivingFormulas(cell => !dataCellDef(cell));
        }



        public void InsertFormulas(ExcelWorksheet worksheet, string[] headers)
        {
            foreach(string header in headers)
            {
                Tuple<int, int> headerCellCoords = FindHeaderCell(worksheet, header);
                AddFormulas(worksheet, headerCellCoords.Item1, headerCellCoords.Item2);
            }
        }




        /// <summary>
        /// Finds the cell that matches the specified header (and is therefore the column that needs formulas)
        /// </summary>
        /// <param name="worksheet">the worksheet in need of formulas</param>
        /// <param name="header">the text that the header cell must match</param>
        /// <returns>the row and column (as a tuple) of the header cell at the top of the formula column</returns>
        private Tuple<int, int> FindHeaderCell(ExcelWorksheet worksheet, string header)
        {
            ExcelIterator iter = new ExcelIterator(worksheet);
            return iter.FindAllMatchingCoordinates(cell => FormulaManager.TextMatches(cell.Text, header)).First();
        }



        /// <summary>
        /// Gives each cell in the specified column a formula if needed
        /// </summary>
        /// <param name="worksheet">the worksheet in need of formulas</param>
        /// <param name="row">the row number of the header of the column getting formulas</param>
        /// <param name="col">the column getting formulas</param>
        private void AddFormulas(ExcelWorksheet worksheet, int row, int col)
        {
            ExcelIterator iter = new ExcelIterator(worksheet, row + 1, col);

            var summaryCells = iter.GetCells(ExcelIterator.SHIFT_DOWN, cell => columnEnds(cell));

            foreach (ExcelRange cell in summaryCells)
            {
                //if this.columnEnds is set to allow empty cells, those empty cells should be skipped
                if (FormulaManager.IsEmptyCell(cell))
                {
                    continue;
                }


                int startColumn = GetFormulaStartColumn(worksheet, cell.Start.Row, col);
                string formula = BuildFormula(worksheet, cell.Start.Row, startColumn, col - 1);

                FormulaManager.PutFormulaInCell(cell, formula);
            }
        }




        /// <summary>
        /// Finds the column number of the first column in this formula range.
        /// </summary>
        /// <param name="worksheet">the worksheet in need of formulas</param>
        /// <param name="row">the row this formula is for</param>
        /// <param name="startCol">the column we should start iterating from</param>
        /// <returns>the column number of the leftmost column in the formula</returns>
        private int GetFormulaStartColumn(ExcelWorksheet worksheet, int row, int startCol)
        {
            ExcelIterator iter = new ExcelIterator(worksheet, row, startCol);
            var lastCell = iter.GetCells(ExcelIterator.SHIFT_LEFT, cell => outsideFormula(cell)).Last();
            return lastCell.End.Column;
        }




        /// <summary>
        /// Builds a formula that spans the horizontal area between the specified start and end columns (inclusive)
        /// </summary>
        /// <param name="worksheet">the worksheet in need of formulas</param>
        /// <param name="row">the row this formula is for</param>
        /// <param name="startCol">the start (leftmost) column of the formula</param>
        /// <param name="endCol">the end (rightmost) column of the formula</param>
        /// <returns>a string with the proper formula to sum up the specified range</returns>
        private string BuildFormula(ExcelWorksheet worksheet, int row, int startCol, int endCol)
        {
            if(startCol > endCol)
            {
                return null; //dont insert a formula
            }

            ExcelRange formulaRange = worksheet.Cells[row, startCol, row, endCol];

            return "SUM(" + formulaRange.Address + ")";
        }



        /// <summary>
        /// Sets when the formula generator stops moving left to find cells to be included in the formula
        /// </summary>
        /// <param name="isOutsideFormula">a function that returns true when the given cell should not be included in the formula</param>
        public void SetOutsideFormulaDefenition(IsOutsideFormula isOutsideFormula)
        {
            this.outsideFormula = isOutsideFormula;
        }



        /// <summary>
        /// Used to modify the way this formula generator determans when to stop putting formulas in a column
        /// </summary>
        /// <param name="endFormulaColumn">a function that returns true when a cell shouldnt get a formula</param>
        public void SetStopGivingFormulas(StopGivingFormulas endFormulaColumn)
        {
            this.columnEnds = endFormulaColumn;
        }



        public void SetDataCellDefenition(IsDataCell isDataCell)
        {
            this.dataCellDef = isDataCell;
        }
    }
}
