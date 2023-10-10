using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelCleanerNet45.FormulaGeneration.ReportSpecificGenerators
{

    /// <summary>
    /// Version of PeriodicFormulaGenerator that is designed to work for ReportCashRecipts report
    /// </summary>
    internal class ReportCashRecipts : PeriodicFormulaGenerator
    {

        /// <inheritdoc/>
        protected override void ProcessFormulaRange(ExcelWorksheet worksheet, ref int row, int dataCol)
        {
            //First get the start and end of the section
            int start = row;
            AdvanceToLastRow(worksheet, ref row);
            int end = row;


            if(start == end) //if there is only one row there
            {
                return; // skip this section
            }


            ProcessInlineFormulas(worksheet, start, end, dataCol);

            ProcessSectionTotals(worksheet, start, end, dataCol);

        }



        /// <summary>
        /// Adds the inline formulas (those that only total up part of the section) to each section.
        /// </summary>
        /// <param name="worksheet">the worksheet in need of formulas</param>
        /// <param name="topOfSection">top row of the formula section currently being worked on</param>
        /// <param name="bottomOfSection">bottom row of the formula section currently being worked on</param>
        /// <param name="dataCol">the column that needs the formulas</param>
        protected virtual void ProcessInlineFormulas(ExcelWorksheet worksheet, int topOfSection, int bottomOfSection, int dataCol)
        {
            ExcelRange cell;
            int formulaStart = topOfSection;
            for (int i = topOfSection; i <= bottomOfSection; i++)
            {
                cell = worksheet.Cells[i, dataCol];

                if (HasTopBorder(cell))
                {
                    cell.Formula = FormulaManager.GenerateFormula(worksheet, formulaStart, i - 1, dataCol);
                    cell.Style.Locked = true;

                    formulaStart = i + 1;
                }
            }
        }




        /// <summary>
        /// Adds formula totals to the "Total" row of each section.
        /// </summary>
        /// <param name="worksheet">the worksheet in need of formulas</param>
        /// <param name="topOfSection">the top row of the formula section</param>
        /// <param name="bottomOfSection">the bottom row of the formula section</param>
        /// <param name="dataCol">the column that needs the formula</param>
        protected virtual void ProcessSectionTotals(ExcelWorksheet worksheet, int topOfSection, int bottomOfSection, int dataCol)
        {
            //first assert that the text "Total:" is found here
            if (!RowHasTotalHeader(worksheet, bottomOfSection))
            {
                return;
            }


            ExcelRange summaryCell = worksheet.Cells[bottomOfSection, dataCol];
            summaryCell.Formula = BuildSectionTotalFormula(worksheet, topOfSection, bottomOfSection, dataCol);
            summaryCell.Style.Locked = true;
        }



        /// <summary>
        /// Checks if the specified row has the header "Total:" in it
        /// </summary>
        /// <param name="worksheet">the worksheet in need of formulas</param>
        /// <param name="row">the row that should be checked for the header</param>
        /// <returns>true if the header is found in the row, and false otherwise</returns>
        protected bool RowHasTotalHeader(ExcelWorksheet worksheet, int row)
        {
            ExcelIterator iter = new ExcelIterator(worksheet, row, 1);

            foreach(ExcelRange cell in iter.GetCells(ExcelIterator.SHIFT_RIGHT))
            {
                if(cell.Text == "Total:")
                {
                    return true;
                }
            }

            return false;
        }



        /// <summary>
        /// Builds a formula that sums all non summary cells between the specified top and bottom of the section
        /// </summary>
        /// <param name="worksheet">the worksheet in need of formulas</param>
        /// <param name="topOfSection">the top row of the formula section</param>
        /// <param name="bottomOfSection">the bottom row of the formula section</param>
        /// <param name="dataCol">the column being summed</param>
        /// <returns>a string formula of the non-summary cells in the column</returns>
        protected virtual string BuildSectionTotalFormula(ExcelWorksheet worksheet, int topOfSection, int bottomOfSection, int dataCol)
        {
            StringBuilder formula = new StringBuilder("SUM(");
            ExcelRange cell;
            for(int i = bottomOfSection - 1; i >= topOfSection; i--)
            {
                cell = worksheet.Cells[i, dataCol];
                if (!FormulaManager.CellHasFormula(cell))
                {
                    formula.Append(cell.Address);
                    formula.Append(",");
                }
            }

            formula.Remove(formula.Length - 1, 1);
            formula.Append(")");
            return formula.ToString();
        }
    }
}
