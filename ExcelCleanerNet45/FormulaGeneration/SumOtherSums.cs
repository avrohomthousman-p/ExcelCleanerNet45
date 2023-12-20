using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelCleanerNet45.FormulaGeneration
{
    /// <summary>
    /// Functions the same way as the full table formula generator except that it only adds the non-formula cells in
    /// the column instead of adding all cells (formula or not).
    /// 
    /// The only argument(s) for this class are the headers that are found at the start of the row with the 
    /// summary formulas.
    /// </summary>
    internal class SumOtherSums : FullTableFormulaGenerator
    {

        public SumOtherSums()
        {
            //in this version we override the defualt behavior set in the superclass
            base.beyondFormulaRange = new IsBeyondFormulaRange(cell => !FormulaManager.IsEmptyCell(cell) && !isDataCell(cell));
        }



        /// <inheritdoc/>
        protected override void FillInFormulas(ExcelWorksheet worksheet, int row, int col)
        {
            ExcelIterator iter = new ExcelIterator(worksheet, row, col + 1);

            foreach (ExcelRange cell in iter.GetCells(ExcelIterator.SHIFT_RIGHT))
            {
                if (HasNoTextOrFormulas(cell))
                {
                    continue;
                }


                //cell.Formula = BuildFormula(worksheet, iter.GetCurrentRow(), iter.GetCurrentCol());
                cell.CreateArrayFormula(BuildFormula(worksheet, iter.GetCurrentRow(), iter.GetCurrentCol()));
                cell.Style.Locked = true;
                cell.Style.Hidden = false;
                cell.Calculate();

                Console.WriteLine("Cell " + cell.Address + " has been given this formula: " + cell.Formula);
            }
        }



        /// <summary>
        /// Checks if the specified cell has niether text nor formulas in it and is therefore empty.
        /// </summary>
        /// <param name="cell">the cell being checked</param>
        /// <returns>true if the cell has no formulas and no text, or false if it has either one.</returns>
        protected bool HasNoTextOrFormulas(ExcelRange cell)
        {
            return FormulaManager.IsEmptyCell(cell) && (cell.Formula == null || cell.Formula.Length == 0);
        }



        /// <summary>
        /// Builds a formula that adds all other formulas in the formula range
        /// </summary>
        /// <param name="worksheet">the worksheet in need of formulas</param>
        /// <param name="headerRow">the row number of the formula header</param>
        /// <param name="headerCol">the column number of the formula header</param>
        /// <returns>a formula to be inserted in the proper cell</returns>
        protected virtual string BuildFormula(ExcelWorksheet worksheet, int headerRow, int headerCol)
        {
            int rangeTop = GetTopCellInRange(worksheet, headerRow, headerCol);
            ExcelRange range = worksheet.Cells[rangeTop, headerCol, headerRow - 1, headerCol];


            //Formula to add up all cells that don't contain a formula
            //The _xlfn fixes a bug in excel
            return $"SUM(IF(_xlfn.ISFORMULA({range.Address}), 0, {range.Address}))";
        }




        /// <summary>
        /// Finds the topmost cell that should be included in the formula range
        /// </summary>
        /// <param name="worksheet">the worksheet in need of formulas</param>
        /// <param name="headerRow">the row number of the formula header (the cell under the last cell in formula range)</param>
        /// <param name="headerCol">the column number of the formula header (the cell under the last cell in formula range)</param>
        /// <returns>the row number of the top of the formula range</returns>
        protected virtual int GetTopCellInRange(ExcelWorksheet worksheet, int headerRow, int headerCol)
        {
            ExcelIterator iter = new ExcelIterator(worksheet, headerRow - 1, headerCol);

            var cells = iter.GetCellCoordinates(ExcelIterator.SHIFT_UP, cell => base.beyondFormulaRange(cell));

            return cells.Last().Item1;
        }
    }
}
