using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace ExcelCleanerNet45.FormulaGeneration
{
    /// <summary>
    /// Override of the RowSegmentFormulaGenerator that sums only the formulas in the range and not all cells.
    /// 
    /// This class is to RowSegmentFormulaGenerator as SumOtherSums is to FullTableFormulaGenerator.
    /// 
    /// Arguments for this class are the same as the RowSegmentFormulaGenerator.
    /// </summary>
    internal class SumOfSumsPeriodic : RowSegmentFormulaGenerator
    {
        protected override void FillInFormulas(ExcelWorksheet worksheet, int startRow, int endRow, int col)
        {
            ExcelRange cell;



            //Often there are multiple columns that require a formula, so we need to iterate
            //and apply the formulas in many columns
            for (col++; col <= worksheet.Dimension.End.Column; col++)
            {
                cell = worksheet.Cells[endRow, col];

                if (this.isDataCell(cell))
                {
                    if (base.trimFormulaRange)
                    {
                        startRow += CountEmptyCellsOnTop(worksheet, startRow, endRow, col); //Skip the whitespace on top
                    }

                    //string formula = FormulaManager.GenerateFormula(worksheet, startRow, endRow - 1, col);
                    //FormulaManager.PutFormulaInCell(cell, formula);
                    cell.CreateArrayFormula(BuildFormula(worksheet, startRow, endRow - 1, col));
                    cell.Style.Locked = true;
                    cell.Style.Hidden = false;
                    cell.Calculate();
                }
                else if (!FormulaManager.IsEmptyCell(cell))
                {
                    return;
                }
            }
        }



        /// <summary>
        /// Builds a formula that adds all other formulas in the formula range
        /// </summary>
        /// <param name="worksheet">the worksheet in need of formulas</param>
        /// <param name="startRow">the topmost row to be included in the formula</param>
        /// <param name="endCol">the bottom-most row to be included in the formula</param>
        /// <param name="col">the column the formula spans</param>
        /// <returns>a formula to be inserted in the proper cell</returns>
        protected virtual string BuildFormula(ExcelWorksheet worksheet, int startRow, int endRow, int col)
        {
            ExcelRange range = worksheet.Cells[startRow, col, endRow, col];


            //Formula to add up all cells that don't contain a formula
            //The _xlfn fixes a bug in excel
            return $"SUM(IF(_xlfn.ISFORMULA({range.Address}), {range.Address}, 0))";
        }
    }
}
