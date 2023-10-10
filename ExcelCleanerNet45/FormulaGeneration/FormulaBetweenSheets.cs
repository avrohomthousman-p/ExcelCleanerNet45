using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelCleanerNet45.FormulaGeneration
{
    /// <summary>
    /// Implementation of IFormulaGenerator that builds formulas that sum up the totals from the bottom of multiple worksheets.
    /// The header argments for this class should simply be all the sheet numbers that are to be included in the sum, expressed 
    /// in the format "sheet[sheetNum]" with no brackets. The sheet numbers are zero based.
    /// 
    /// Note: the formula created by this implementation is always placed at the bottom of the current worksheet 
    /// and always gets its data from the bottom row of other worksheets. This class does not provide an option
    /// for changing where in a worksheet the data will come from.
    /// </summary>
    internal class FormulaBetweenSheets : IFormulaGenerator
    {
        private IsDataCell dataCellDef = new IsDataCell(
            cell => FormulaManager.IsDollarValue(cell) || FormulaManager.CellHasFormula(cell));



        public void InsertFormulas(ExcelWorksheet mainWorksheet, string[] headers)
        {
            bool[] isNegative = headers.Select(s => s.StartsWith("-")).ToArray();
            int[] sheetsToAdd = headers.Select(s =>
            {
                if(s.StartsWith("-"))
                {
                    return Int32.Parse(s.Substring(6)); //-sheet3 => 3
                }
                else
                {
                    return Int32.Parse(s.Substring(5)); //sheet3 => 3
                }
            })
            .ToArray();




            ExcelWorkbook workbook = mainWorksheet.Workbook;

            int summaryCellNum = 0; //tracks which summary cell we are currently working on
            ExcelIterator mainIterator = new ExcelIterator(mainWorksheet, mainWorksheet.Dimension.End.Row, 1);
            foreach(ExcelRange cell in mainIterator.GetCells(ExcelIterator.SHIFT_RIGHT))
            {
                if (!dataCellDef(cell))
                {
                    continue;
                }

                summaryCellNum++;
                cell.Formula = BuildFormula(workbook, sheetsToAdd, isNegative, summaryCellNum, mainWorksheet.Index);
                cell.Style.Locked = true;

                Console.WriteLine("Cell " + cell.Address + " has been given this formula: " + cell.Formula);
            }
        }



        /// <summary>
        /// Builds the formula needed for the current worksheet.
        /// </summary>
        /// <param name="workbook">the workbook that needs formulas</param>
        /// <param name="sheets">the worksheet indexes that should be included in the formula</param>
        /// <param name="isNegative">an array of bools telling you which worksheets should be subtracted instead of added</param>
        /// <param name="summaryCell">the data cell that we should include in the formula</param>
        /// <param name="mainWorksheet">the index of the worksheet that gets the formula</param>
        /// <returns>a formula that can be used to add up the correct cells</returns>
        private string BuildFormula(ExcelWorkbook workbook, int[] sheets, bool[] isNegative, int summaryCell, int mainWorksheet)
        {
            StringBuilder formula = new StringBuilder("SUM(");


            ExcelWorksheet currentWorksheet;
            bool isMainWorksheet;
            for(int i = 0; i < sheets.Length; i++)
            {
                currentWorksheet = workbook.Worksheets[sheets[i]];
                isMainWorksheet = (sheets[i] == mainWorksheet);
                
                string address = GetCellFromOtherWorksheet(currentWorksheet, summaryCell, isMainWorksheet);
                if(address == null)
                {
                    Console.WriteLine($"Warning: sheet {sheets[i]} does not contain data cells and was skipped");
                    continue;
                }


                if (isNegative[i])
                {
                    formula.Append("-");
                }

                if (!isMainWorksheet)
                {
                    formula.Append("Sheet" + (sheets[i] + 1) + "!");
                }

                formula.Append(address);
                formula.Append(",");
            }


            formula.Remove(formula.Length - 1, 1);
            formula.Append(")");
            return formula.ToString();
        }



        /// <summary>
        /// Gets the address of the summary cell in the specified worksheet.
        /// </summary>
        /// <param name="worksheet">the worksheet where the desired summary cell can be found</param>
        /// <param name="summaryCell">which cell we should return (e.g. 5 would mean we return the 5th data cell)</param>
        /// <param name="isMainWorksheet">true if the specified worksheet is also the main worksheet, and false otherwise</param>
        /// <returns>the address of the cell that should be included in the formula, or null if no appropriate cell is found</returns>
        private string GetCellFromOtherWorksheet(ExcelWorksheet worksheet, int summaryCell, bool isMainWorksheet)
        {
            int row = worksheet.Dimension.End.Row;

            //on the main worksheet we sum the second to last row instead of the last (to avoid circular formulas)
            if (isMainWorksheet)
            {
                row = FindNextDataRow(worksheet);
            }

            int summaryCellsFound = 0;
            ExcelIterator iter = new ExcelIterator(worksheet, row, 1);
            foreach(ExcelRange cell in iter.GetCells(ExcelIterator.SHIFT_RIGHT))
            {
                if (!dataCellDef(cell))
                {
                    continue;
                }


                summaryCellsFound++;
                if(summaryCellsFound == summaryCell)
                {
                    return cell.Address;
                }
            }


            return null;
        }



        /// <summary>
        /// Finds the last row in the worksheet that has data cells in it (starting from the second to last row of the worksheet)
        /// </summary>
        /// <param name="worksheet">the worksheet containing our formula data</param>
        /// <returns>the row the formula data is on</returns>
        private int FindNextDataRow(ExcelWorksheet worksheet)
        {
            ExcelIterator iter = new ExcelIterator(worksheet, worksheet.Dimension.End.Row - 1, worksheet.Dimension.End.Column);

            return iter.FindAllCellsReverse().First(cell => dataCellDef(cell)).End.Row;
        }



        public void SetDataCellDefenition(IsDataCell isDataCell)
        {
            this.dataCellDef = isDataCell;
        }
    }
}
