using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelCleanerNet45.FormulaGeneration
{

    /// <summary>
    /// An override of RowSegmentFormulaGenerator that adds only the formulas in the range, and not the regular 
    /// data cells.
    /// 
    /// Header arguments for this class should be passed in in the same manner as the RowSegmentFormulGenerator.
    /// </summary>
    internal class SumWithinSegmentGenerator : RowSegmentFormulaGenerator
    {

        private bool useArrayFormula = true;
        /// <summary>
        /// controls what kind of formula is used, and Array formula of a regular formula.
        /// 
        /// For large reports, Array Formulas are reccomended.
        /// </summary>
        public bool UseArrayFormula
        {
            get { return useArrayFormula; }
            set { useArrayFormula = value; }
        }



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

                    if (useArrayFormula)
                    {
                        ExcelRange range = worksheet.Cells[startRow, col, endRow - 1, col];
                        cell.CreateArrayFormula(BuildFormula(range));
                    }
                    else
                    {
                        cell.Formula = BuildNonArrayFormula(worksheet, startRow, endRow - 1, col);
                    }
                    
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
        /// Builds an array formula that adds all other formulas in the formula range
        /// </summary>
        /// <param name="range">the range the formula must span</param>
        /// <returns>a formula to be inserted in the proper cell</returns>
        protected virtual string BuildFormula(ExcelRange range)
        {
            //Formula to add up all cells that don't contain a formula
            //The _xlfn fixes a bug in excel
            return $"SUM(IF(_xlfn.ISFORMULA({range.Address}), 0, {range.Address}))";
        }



        /// <summary>
        /// Builds a regular (non array) formula that adds up all other formulas in the range
        /// </summary>
        /// <param name="worksheet">the worksheet in need of formulas</param>
        /// <param name="topRow">the topmost row to be included in the formula range</param>
        /// <param name="bottomRow">the bootom-most row to be included in the formula range</param>
        /// <param name="col">the column the formula range is in</param>
        /// <returns>a formula that can be used as the sum of all formulas in the range</returns>
        protected virtual string BuildNonArrayFormula(ExcelWorksheet worksheet, int topRow, int bottomRow, int col)
        {
            List<string> formulaCells = GetFormulaCellsInRange(worksheet, topRow, bottomRow, col);

            StringBuilder result = new StringBuilder("SUM(");

            foreach(string address in formulaCells)
            {
                result.Append(address);
                result.Append(",");
            }


            result.Remove(result.Length - 1, 1); //remove final comma
            result.Append(")");

            return result.ToString();
        }




        /// <summary>
        /// Gets a list of addresses of each cell in the specified range that should be included in the formula
        /// </summary>
        /// <param name="worksheet">the worksheet being given formulas</param>
        /// <param name="topRow">the top-most row that is part of the formula range</param>
        /// <param name="bottomRow">the bottom-most row that is part of the formula range</param>
        /// <param name="col">the column the formula should cover</param>
        /// <returns>a list of addresses that should be added up by the formula</returns>
        protected virtual List<string> GetFormulaCellsInRange(ExcelWorksheet worksheet, int topRow, int bottomRow, int col)
        {
            List<string> addresses = new List<string>();

            ExcelRange cell;

            for(int row = topRow; row <= bottomRow; row++)
            {
                cell = worksheet.Cells[row, col];

                if (!FormulaManager.CellHasFormula(cell) && isDataCell(cell))
                {
                    addresses.Add(cell.Address);
                }
            }

            return addresses;
        }
    }
}
