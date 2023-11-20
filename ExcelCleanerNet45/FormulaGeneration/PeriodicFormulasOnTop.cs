using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml;

namespace ExcelCleanerNet45.FormulaGeneration
{

    public delegate bool IsSummaryCell(ExcelRange cell);
    public delegate bool IsKeyOfSection(ExcelRange cell);




    /// <summary>
    /// Implementation of IFormulaGenerator that adds formulas to the top of "sections" found inside data columns
    /// of the worksheet. A section is defined as a series of data cells that all corrispond to a single "key" which 
    /// appears on the top left of that section. As an example of this, look at the report VendorInvoiceReport.
    /// 
    /// This class is a lot like the PeriodicFormulaGenerator, but it puts the formulas on the top of each section
    /// instead of the bottom.
    /// 
    /// The string arguments for this class should be the titles of each data column that needs
    /// formulas (meaning, the header text that can be found at the top of each column that needs formulas).
    /// </summary>
    internal class PeriodicFormulasOnTop : IFormulaGenerator
    {
        private IsDataCell isDataCell = new IsDataCell(FormulaManager.IsDollarValue); //default implementation
        private IsSummaryCell isSummaryCell = new IsSummaryCell(
            (cell => FormulaManager.IsDollarValue(cell) && cell.Style.Font.Bold));    //defualt implementation

        private IsKeyOfSection cellHasKey = new IsKeyOfSection(
            (cell => !FormulaManager.IsEmptyCell(cell) 
                    && !FormulaManager.IsDollarValue(cell) 
                    && cell.Style.Font.Bold));                                        //default implementation



        protected int firstRowOfTable = -1;




        public virtual void SetDataCellDefenition(IsDataCell isDataCell)
        {
            this.isDataCell = isDataCell;
        }



        public virtual void SetSummaryCellDefenition(IsSummaryCell summaryCellDef)
        {
            this.isSummaryCell = summaryCellDef;
        }




        public virtual void SetSectionKeyDefenition(IsKeyOfSection predicate)
        {
            this.cellHasKey = predicate;
        }




        public void InsertFormulas(ExcelWorksheet worksheet, string[] headers)
        {
            if(headers.Length == 0)
            {
                return;
            }


            FindStartOfTable(worksheet, headers[0]);

            List<int> keys = FindAllSectionKeys(worksheet);

            List<int> dataColumns = FindAllDataColumns(worksheet, headers);


            //now add a formula for each key in the appropriate columns
            int startRow;
            int endRow;
            ExcelRange summaryCell;
            for(int i = 0; i < keys.Count - 1; i++)
            {
                startRow = keys[i] + 1; //formula should start from the row after the summary cell
                endRow = keys[i + 1] - 1; //formula should go until (but not including) the next formula cell

                foreach(int col in dataColumns)
                {
                    summaryCell = worksheet.Cells[keys[i], col];

                    string formula = FormulaManager.GenerateFormula(worksheet, startRow, endRow, col);

                    FormulaManager.PutFormulaInCell(summaryCell, formula);
                }
            }

        }



        /// <summary>
        /// Finds the first row of the worksheet that is actually part of the table. Anything above that row is
        /// considered a major header, and anything on it or below is either data or a minor header.
        /// </summary>
        /// <param name="worksheet">the worksheet in need of formulas</param>
        /// <param name="header">the header to look for to signal the start of the table</param>
        protected virtual void FindStartOfTable(ExcelWorksheet worksheet, string header)
        {
            ExcelIterator iter = new ExcelIterator(worksheet);

            foreach(ExcelRange cell in iter.FindAllCells())
            {
                if(FormulaManager.TextMatches(cell.Text, header))
                {
                    firstRowOfTable = iter.GetCurrentRow();
                    return;
                }
            }


            throw new Exception("Formulas could not be added becuase no data was found in the worksheet.");
        }




        /// <summary>
        /// Builds a list containing the row numbers where each key can be found. The last entry will be 
        /// a reference to a row that is outside the table, indicating that the section before it spans 
        /// untill the end of the table.
        /// </summary>
        /// <param name="worksheet">the worksheet in need of formulas</param>
        /// <returns>a list of all rows where keys can be found</returns>
        protected virtual List<int> FindAllSectionKeys(ExcelWorksheet worksheet)
        {
            ExcelIterator iter = new ExcelIterator(worksheet, firstRowOfTable + 1, 1);


            //find all the cells with keys
            List<int> keys = iter.GetCells(ExcelIterator.SHIFT_DOWN)
                        .Where(cell => cellHasKey(cell))
                        .Select(cell => cell.Start.Row)
                        .ToList();




            //add a row number to represent the end of the last section
            if (WorksheetHasFinalSummary(worksheet))
            {
                keys.Add(worksheet.Dimension.End.Row);
            }
            else
            {
                keys.Add(worksheet.Dimension.End.Row + 1);
            }
             

            return keys;
        }



        /// <summary>
        /// Checks if the worksheet has a full table summary row at the bottom. Such a row should
        /// not be included in the formula for the bottom-most section.
        /// </summary>
        /// <param name="worksheet">the worksheet getting the formulas</param>
        /// <returns>true if the worksheet has a final summary row at the bottom and false otherwise</returns>
        protected virtual bool WorksheetHasFinalSummary(ExcelWorksheet worksheet)
        {
            ExcelIterator iter = new ExcelIterator(worksheet, worksheet.Dimension.End.Row, 1);

            return iter.GetCells(ExcelIterator.SHIFT_RIGHT).Any(cell => isSummaryCell(cell));
        }




        /// <summary>
        /// Finds all the columns in the table that require formulas in them
        /// </summary>
        /// <param name="worksheet">the worksheet in need of formulas</param>
        /// <param name="headers">all the headers that will be found at the top of a column that needs formulas</param>
        /// <returns>a list of all the columns that require formulas</returns>
        protected virtual List<int> FindAllDataColumns(ExcelWorksheet worksheet, string[] headers)
        {
            ExcelIterator iter = new ExcelIterator(worksheet, firstRowOfTable, 1);



            Func<ExcelRange, bool> matchesAHeader = cell => 
            { 
                foreach(string header in headers)
                {
                    if(FormulaManager.TextMatches(cell.Text, header))
                    {
                        return true;
                    }
                }

                return false;
            };



            //Get all columns with text matching one of the headers (and therefore need formulas)
            return iter.GetCells(ExcelIterator.SHIFT_RIGHT)
                        .Where(matchesAHeader)
                        .Select(cell => cell.Start.Column)
                        .ToList();
        }
    }
}
