using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelCleanerNet45.FormulaGeneration
{

    /// <summary>
    /// An override of the FullTableFormulaGenerator that adds up all the bold cells only. This is 
    /// usefull for any report that needs to add up only all the bold cells in the range (like TrialBalance).
    /// 
    /// Note: if the report is very very long, the formula might fail becuase it went over the maximum number
    /// of characters allowed in an excel formula.
    /// </summary>
    internal class SumOnlyBolds : FullTableFormulaGenerator
    {


        //this is the same code as the parent class except that it changes how the formula is built

        /// <inheritdoc/>
        protected override void FillInFormulas(ExcelWorksheet worksheet, int row, int col)
        {
            iter.SetCurrentLocation(row, col);

            foreach (ExcelRange cell in iter.GetCells(ExcelIterator.SHIFT_RIGHT))
            {
                if (FormulaManager.IsEmptyCell(cell) || !isDataCell(cell))
                {
                    continue;
                }


                int topRowOfRange = FindTopRowOfFormulaRange(worksheet, row, col);

                string formula = GenerateFormula(worksheet, topRowOfRange, row - 1, iter.GetCurrentCol());

                FormulaManager.PutFormulaInCell(cell, formula, false);
            }

        }



        /// <summary>
        /// Builds a formula that adds all bold cells in the specified range
        /// </summary>
        /// <param name="worksheet">the worksheet getting the formula</param>
        /// <param name="topRow">the top row of the range</param>
        /// <param name="bottomRow">the bottom row of the range</param>
        /// <param name="col">the column the range is in</param>
        /// <returns>the formula to be used in the worksheet</returns>
        private string GenerateFormula(ExcelWorksheet worksheet, int topRow, int bottomRow, int col)
        {
            StringBuilder formula = new StringBuilder("SUM(");

            ExcelRange cell;
            for(int i = topRow; i <= bottomRow; i++)
            {
                cell = worksheet.Cells[i, col];
                if (cell.Style.Font.Bold)
                {
                    formula.Append(cell.Address);
                    formula.Append(",");
                }
            }

            //remove last comma
            formula.Remove(formula.Length - 1, 1);

            formula.Append(")");

            return formula.ToString();
        }
    }
}
