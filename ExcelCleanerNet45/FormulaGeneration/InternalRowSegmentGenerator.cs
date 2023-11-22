using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelCleanerNet45.FormulaGeneration
{
    /// <summary>
    /// An override of the RowSegmentFormulaGenerator that only adds formulas to row segments inside other row segments.
    /// This Formula Generator is not being used at the momnet. It was built due to a misunderstanding in how a report works
    /// 
    /// Arguments for this formula generator should be the same as the arguments required by the RowSegmentFormulaGenerator
    /// class.
    /// </summary>
    internal class InternalRowSegmentGenerator : RowSegmentFormulaGenerator
    {


        protected override IEnumerable<Tuple<int, int, int>> GetRowRangeForFormula(ExcelWorksheet worksheet, string startHeader, string endHeader)
        {
            ExcelIterator iter = new ExcelIterator(worksheet);


            //mark the column containing the headers for the outer section that we are not giving formulas to
            int firstColumn = FindColumnOfOuterSections(worksheet, ref iter, startHeader);
            if(firstColumn == -1)
            {
                yield break; //could not find any instances of this header
            }



            //Now find each inner section start header
            Predicate<ExcelRange> matchesStartHeader = cell => cell.Start.Column > firstColumn 
                                                                && FormulaManager.TextMatches(cell.Text, startHeader);

            var matchingCells = iter.FindAllMatchingCells(matchesStartHeader);



            foreach (ExcelRange cell in matchingCells)
            {
                int endRow = RowSegmentFormulaGenerator.FindEndOfFormulaRange(worksheet, cell.Start.Row, cell.Start.Column, endHeader);

                if (endRow > 0)
                {
                    yield return new Tuple<int, int, int>(cell.Start.Row, endRow, cell.Start.Column);


                    //if we are not on the last row (index out of bounds check)
                    if (endRow < worksheet.Dimension.End.Row)
                    {
                        iter.SetCurrentLocation(endRow + 1, firstColumn + 1); //skip to the the next row (we dont expect 2 headers on one row)
                    }
                }
            }


        }




        /// <summary>
        /// Finds the column that contains the headers of the outer sections (which get no formulas).
        /// </summary>
        /// <param name="worksheet">the worksheet in need of formulas</param>
        /// <param name="iter">a reference to the iterator we should use</param>
        /// <param name="startHeader">the text of the header we are looking for</param>
        /// <returns>the column number of the column containing the headers of the outer sections</returns>
        protected virtual int FindColumnOfOuterSections(ExcelWorksheet worksheet, ref ExcelIterator iter, string startHeader)
        {
            ExcelRange firstCell = iter.GetFirstMatchingCell(cell => FormulaManager.TextMatches(cell.Text, startHeader));
            iter.SkipRow(); //ensure the calling method continues iteration from the column after the one with the header

            if (firstCell == null)
            {
                return -1;
            }
            else
            {
                return firstCell.Start.Column;
            }
        }
    }
}
