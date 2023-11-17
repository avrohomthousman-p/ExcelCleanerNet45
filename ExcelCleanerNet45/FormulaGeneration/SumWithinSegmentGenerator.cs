using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelCleanerNet45.FormulaGeneration
{

    /// <summary>
    /// An override of RowSegmentFormulaGenerator that adds only the formulas in the range, and not the regular data cells.
    /// 
    /// Header arguments for this class should be passed in in the same manner as the RowSegmentFormulGenerator.
    /// </summary>
    internal class SumWithinSegmentGenerator : RowSegmentFormulaGenerator
    {
        /// <inheritdoc/>
        protected override void FillInFormulas(ExcelWorksheet worksheet, int startRow, int endRow, int col)
        {
            ExcelRange cell;



            //Often there are multiple columns that require a formula, so we need to iterate
            //and apply the formulas in many columns
            for (col++; col <= worksheet.Dimension.End.Column; col++)
            {
                cell = worksheet.Cells[endRow, col];

                if (base.isDataCell(cell))
                {
                    if (base.trimFormulaRange)
                    {
                        startRow += CountEmptyCellsOnTop(worksheet, startRow, endRow, col); //Skip the whitespace on top
                    }


                    //this is the part that differs from the parent class
                    ExcelRange range = worksheet.Cells[startRow, col, endRow - 1, col];

                    cell.CreateArrayFormula(BuildFormula(range));
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
        /// <param name="range">the range the formula must span</param>
        /// <returns>a formula to be inserted in the proper cell</returns>
        protected virtual string BuildFormula(ExcelRange range)
        {
            //Formula to add up all cells that don't contain a formula
            //The _xlfn fixes a bug in excel
            return $"SUM(IF(_xlfn.ISFORMULA({range.Address}), 0, {range.Address}))";
        }
    }
}
