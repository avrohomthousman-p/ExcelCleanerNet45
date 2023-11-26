using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelCleanerNet45.FormulaGeneration
{
    /// <summary>
    /// Extentsion of the RowSegmentFormulaGenerator that continues adding formulas after hitting a cell with
    /// a percentage in it.
    /// 
    /// This class also fixes a bug in the ProfitAndLossBudget report that puts the YTD header in the wrong column.
    /// 
    /// The string arguments from this class are the same as the parent class (RowSegmentFormulaGenerator).
    /// </summary>
    internal class ProfitAndLossBudgetFormulas : RowSegmentFormulaGenerator
    {

        public override void InsertFormulas(ExcelWorksheet worksheet, string[] headers)
        {
            base.InsertFormulas(worksheet, headers);

            FixYTDHeader(worksheet);
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

                if (this.isDataCell(cell))
                {
                    if (base.trimFormulaRange)
                    {
                        startRow += CountEmptyCellsOnTop(worksheet, startRow, endRow, col); //Skip the whitespace on top
                    }
                    string formula = FormulaManager.GenerateFormula(worksheet, startRow, endRow - 1, col);

                    FormulaManager.PutFormulaInCell(cell, formula);
                }
                else if (FormulaManager.IsPercentage(cell))
                {
                    continue; //dont put a formula here, but dont stop inserting formulas in this row
                }
                else if (!FormulaManager.IsEmptyCell(cell))
                {
                    return;
                }
            }

        }



        /// <summary>
        /// Moves the YTD header to the nearest data cell if needed
        /// </summary>
        /// <param name="worksheet">the worksheet being cleaned</param>
        protected virtual void FixYTDHeader(ExcelWorksheet worksheet)
        {
            ExcelIterator iter = new ExcelIterator(worksheet);
            ExcelRange sourceCell = iter.GetFirstMatchingCell(cell => cell.Text == "YTD");

            //if its already in the right place
            if (!IsPercentageColumn(worksheet, sourceCell.Start.Column))
            {
                return;
            }


            int targetColumn = GetDestinationColumn(worksheet, sourceCell.Start.Column);
            if(targetColumn == -1) //no suitable destination column found
            {
                return;
            }



            //copy the data
            ExcelRange destination = worksheet.Cells[sourceCell.Start.Row, targetColumn];

            sourceCell.CopyStyles(destination);
            destination.Value = sourceCell.Value;
            sourceCell.Value = null;
        }




        /// <summary>
        /// Checks if the specified column has at least one percentage value in it.
        /// </summary>
        /// <param name="worksheet">the worksheet being cleaned</param>
        /// <param name="column">the column we are searching for percentages</param>
        /// <returns>true if the column has at least one percentage in it, or false otherwise</returns>
        protected virtual bool IsPercentageColumn(ExcelWorksheet worksheet, int column)
        {
            ExcelIterator iter = new ExcelIterator(worksheet, 1, column);

            return iter.GetCells(ExcelIterator.SHIFT_DOWN).Any(cell => FormulaManager.IsPercentage(cell));
        }




        /// <summary>
        /// Checks if the specified column has at least one dollar value in it.
        /// </summary>
        /// <param name="worksheet">the worksheet being cleaned</param>
        /// <param name="column">the column we are searching for dollar values</param>
        /// <returns>true if the column has at least one dollar in it, or false otherwise</returns>
        protected virtual bool IsDollarColumn(ExcelWorksheet worksheet, int column)
        {
            ExcelIterator iter = new ExcelIterator(worksheet, 1, column);

            return iter.GetCells(ExcelIterator.SHIFT_DOWN).Any(cell => FormulaManager.IsDollarValue(cell));
        }




        /// <summary>
        /// Gets the next column to the right of the source column, that has at leaset one dollar value. If no column
        /// is found, -1 is returned.
        /// </summary>
        /// <param name="worksheet">the worksheet currently being cleaned</param>
        /// <param name="sourceColumn">the column we need to start our search at (after) </param>
        /// <returns>
        /// the column number of the closest dollar column to the right of the source column, or 
        /// -1 if no such column is found
        /// </returns>
        protected virtual int GetDestinationColumn(ExcelWorksheet worksheet, int sourceColumn)
        {
            int targetColumn = sourceColumn + 1;

            while (targetColumn <= worksheet.Dimension.End.Column)
            {
                if (IsDollarColumn(worksheet, targetColumn))
                {
                    break;
                }

                targetColumn++;
            }



            if(targetColumn == worksheet.Dimension.End.Column + 1)
            {
                return -1;
            }
            else
            {
                return targetColumn;
            }
        }
    }
}
