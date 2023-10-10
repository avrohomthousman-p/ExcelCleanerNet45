using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelCleanerNet45.FormulaGeneration.ReportSpecificGenerators
{
    /// <summary>
    /// Implementation of IFormulaGenerator that is specifically designed to clean the report 
    /// VendorInvoiceReportWithJournelAccounts. It should be passed an array of headers each corrisponding
    /// to the headers above the summary cells before the report data.
    /// </summary>
    internal class VendorInvoiceReportFormulas : IFormulaGenerator
    {

        private IsDataCell dataCellDef = new IsDataCell(FormulaManager.IsDollarValue);


        private int firstDataRow;


        public void SetDataCellDefenition(IsDataCell isDataCell)
        {
            this.dataCellDef = isDataCell;
        }



        public void InsertFormulas(ExcelWorksheet worksheet, string[] headers)
        {
            firstDataRow = FindFirstDataRow(worksheet);

            AddFormulasForInvoiceTotals(worksheet);

            foreach(string header in headers)
            {
                AddSummaryFormulas(worksheet, header);
            }
        }



        /// <summary>
        /// Finds the first row of the worksheet that can be considered part of the table (and should be included in 
        /// formulas).
        /// </summary>
        /// <param name="worksheet">the worksheet in need of formulas</param>
        /// <returns>the row number of the topmost row that has data for formulas in it</returns>
        private int FindFirstDataRow(ExcelWorksheet worksheet)
        {

            for (int row = 1; row <= worksheet.Dimension.End.Row; row++)
            {
                if (HasManyEntries(worksheet, row))
                {

                    return row + 1; //we want the first row with actual data, so 1 below the headers

                }

            }

            return 1;
        }



        /// <summary>
        /// Checks if the specified row has at least 5 non-empty cells in it
        /// </summary>
        /// <param name="worksheet">the worksheet in need of formulas</param>
        /// <param name="row">the row to be checked</param>
        /// <returns>true if the specified row contains at least 5 non-empty cells</returns>
        private bool HasManyEntries(ExcelWorksheet worksheet, int row)
        {
            ExcelIterator iter = new ExcelIterator(worksheet, row, 1);

            return iter.GetCells(ExcelIterator.SHIFT_RIGHT).Count(cell => !FormulaManager.IsEmptyCell(cell)) >= 5;
        }



        /// <summary>
        /// Finds all cells in the report that contain the text "Invoice Total: [some amount]" and 
        /// breaks the text and number into two seperate cells, and adds a formula.
        /// </summary>
        /// <param name="worksheet">the worksheet in need of formulas</param>
        private void AddFormulasForInvoiceTotals(ExcelWorksheet worksheet)
        {
            string regex = "Invoice Total: \\$\\d{1,3}(,\\d{3})*[.]\\d\\d";


            //Invoice totals are always in the same column, so we need to move to that column
            ExcelIterator iter = new ExcelIterator(worksheet);
            iter.GetFirstMatchingCell(c => FormulaManager.TextMatches(c.Text, regex));


            //move iterator up 1 to ensure the coming loop doesnt miss the first match
            iter.SetCurrentLocation(iter.GetCurrentRow() - 1, iter.GetCurrentCol());


            var cells = iter.GetCellCoordinates(ExcelIterator.SHIFT_DOWN);
            ExcelRange cell;
            foreach (Tuple<int, int> position in cells)
            {
                cell = worksheet.Cells[position.Item1, position.Item2];

                if(FormulaManager.TextMatches(cell.Text, regex))
                {
                    ExcelRange summaryCell = SplitCell(worksheet, cell);

                    if(summaryCell != null)
                    {
                        summaryCell.Formula = BuildSectionFormula(worksheet, summaryCell.Start.Row, summaryCell.Start.Column + 1);
                        summaryCell.Style.Locked = true;
                        summaryCell.Style.Numberformat.Format = "$#,##0.00;($#,##0.00)";
                        Console.WriteLine("Cell " + cell.Address + " has been given this formula: " + cell.Formula);
                    }
                }
            }
        }



        /// <summary>
        /// Splits the specified cell into two cells, one for the text and one for the data.
        /// </summary>
        /// <param name="worksheet">the worksheet in need of forlumas</param>
        /// <param name="currentLocation">the cell that currently has the text</param>
        /// <returns>
        /// the cell that is now the data part of the original cell and needs a formula, 
        /// or null if the cell could not be split (none of the adjacent cells are availible)
        /// </returns>
        private ExcelRange SplitCell(ExcelWorksheet worksheet, ExcelRange currentLocation)
        {
            string text = currentLocation.Text.Substring(0, currentLocation.Text.IndexOf('$')).Trim();

            int currentRow = currentLocation.Start.Row;
            int currentCol = currentLocation.Start.Column;


            if (currentCol > 1)
            {
                ExcelRange cellToTheLeft = worksheet.Cells[currentRow, currentCol - 1];
                if (FormulaManager.IsEmptyCell(cellToTheLeft))
                {
                    //move the first half of the text over
                    cellToTheLeft.SetCellValue(0, 0, text);
                    currentLocation.CopyStyles(cellToTheLeft);
                    return currentLocation;
                }
            }


            ExcelRange cellToTheRight = worksheet.Cells[currentRow, currentCol + 1];
            if (FormulaManager.IsEmptyCell(cellToTheRight))
            {
                //remove the data part of the text from the current cell
                currentLocation.SetCellValue(0, 0, text);
                return cellToTheRight;
            }


            return null;
        }



        /// <summary>
        /// Finds the area that needs to be included in the formula and builds and returns the formula
        /// </summary>
        /// <param name="worksheet">the worksheet in need of formulas</param>
        /// <param name="row">the row number of the summary cell</param>
        /// <param name="col">the column number of the formula cell</param>
        /// <returns>the formula that should be added to the formula cell</returns>
        private string BuildSectionFormula(ExcelWorksheet worksheet, int row, int col)
        {
            col -= 2; //not entirely sure why, but this is nessecary
            row--; //we want to start above the summary cell (first data cell)

            int bottom = row;
            int top;


            //Iterate upward untill we find a cell with non data text OR we go above the top row
            ExcelRange cell;
            for(; row >= firstDataRow; row--)
            {
                cell = worksheet.Cells[row, col];
                if(!FormulaManager.IsEmptyCell(cell) && !FormulaManager.IsDollarValue(cell))
                {
                    break;
                }
            }

            top = row + 1; //we stop when we hit a non-data cell that ISN'T part of the formula
            cell = worksheet.Cells[top, col, bottom, col];

            return "SUM(" + cell.Address + ")";
        }



        /// <summary>
        /// Adds summary formulas to the worksheet
        /// </summary>
        /// <param name="worksheet">the worksheet in need of formulas</param>
        /// <param name="header">the header we are looking for</param>
        private void AddSummaryFormulas(ExcelWorksheet worksheet, string header)
        {
            ExcelIterator iter = new ExcelIterator(worksheet);
            ExcelRange topCell = iter.GetFirstMatchingCell(c => FormulaManager.TextMatches(c.Text, header));
            int col = topCell.Start.Column;

            ExcelRange topSummaryCell = worksheet.Cells[topCell.End.Row + 1, col];
            ExcelRange bottomSummaryCell = worksheet.Cells[worksheet.Dimension.End.Row, col];

            string formula = BuildSummaryFormula(worksheet, col);

            topSummaryCell.Formula = formula;
            topSummaryCell.Style.Locked = true;

            bottomSummaryCell.Formula = formula;
            bottomSummaryCell.Style.Locked = true;

        }



        /// <summary>
        /// Builds a formula for the summary cells in the report
        /// </summary>
        /// <param name="worksheet">the worksheet in need of formulas</param>
        /// <param name="col">the column number the formula should cover</param>
        /// <returns>a string containing the excel formula needed</returns>
        private string BuildSummaryFormula(ExcelWorksheet worksheet, int col)
        {
            ExcelIterator iter = new ExcelIterator(worksheet, firstDataRow, col);
            StringBuilder formula = new StringBuilder("SUM(");
            foreach (ExcelRange cell in iter.GetCells(ExcelIterator.SHIFT_DOWN))
            {
                //skip the last row
                if (iter.GetCurrentRow() == worksheet.Dimension.End.Row - 1)
                {
                    break;
                }

                if (!FormulaManager.CellHasFormula(cell) && (FormulaManager.IsDollarValue(cell) || FormulaManager.IsEmptyCell(cell)))
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
