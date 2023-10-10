using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelCleanerNet45.FormulaGeneration.ReportSpecificGenerators
{

    /// <summary>
    /// Implementation of IFormulaGenerator that adds a final column and final row with formulas for the totals
    /// of the full worksheet.
    /// </summary>
    internal class RentRollHistorySheet1 : IFormulaGenerator
    {
        private IsDataCell dataCellDef = new IsDataCell(cell => FormulaManager.IsDollarValue(cell));






        public void InsertFormulas(ExcelWorksheet worksheet, string[] headers)
        {
            int headerRow = FindHeaderRow(worksheet);

            AddSummaryRow(worksheet, headerRow);

            AddSummaryColumn(worksheet, headerRow);


            //now call other formula generators to replace our sample data with formulas
            IFormulaGenerator generator = new FullTableSummaryColumn();
            generator.InsertFormulas(worksheet, new string[] { "Total:" });


            generator = new FullTableFormulaGenerator();
            generator.InsertFormulas(worksheet, new string[] { "Total:" });
        }




        /// <summary>
        /// Finds the row in the worksheet that is considered the "header row" - the row with the headers for each data
        /// column. In this implementation, the header row is the first row that has at least 3 non empty cells.
        /// </summary>
        /// <param name="worksheet">the worksheet in need of formulas</param>
        /// <returns>the row number of the header row for this worksheet</returns>
        private int FindHeaderRow(ExcelWorksheet worksheet)
        {

            for(int row = 1; row <= worksheet.Dimension.End.Row; row++)
            {
                int nonEmptyCells = 0;

                for(int col = 1; col <= worksheet.Dimension.End.Column; col++)
                {
                    ExcelRange cell = worksheet.Cells[row, col];

                    if (!FormulaManager.IsEmptyCell(cell))
                    {
                        nonEmptyCells++;
                        if(nonEmptyCells >= 3)
                        {
                            return row;
                        }
                    }
                }
            }


            return 1; //default to first row
        }



        /// <summary>
        /// Adds the new summary row to the end of the worksheet, gives it a header and copys over all the styles
        /// </summary>
        /// <param name="worksheet">the worksheet in need of formulas</param>
        /// <param name="headerRow">the row number of where the new column should have its header</param>
        private void AddSummaryRow(ExcelWorksheet worksheet, int headerRow)
        {
            int summaryRow = worksheet.Dimension.End.Row;

            //add a row that should be the second to last
            worksheet.InsertRow(summaryRow, 1);


            //Put a header at the start of the new row
            ExcelRange headerCell = worksheet.Cells[summaryRow, 1];
            headerCell.SetCellValue(0, 0, "Total:");



            //Fill cells with sample data
            ExcelRange cell;
            ExcelIterator iter = new ExcelIterator(worksheet, headerRow, 3);
            foreach(Tuple<int, int> coords in iter.GetCellCoordinates(ExcelIterator.SHIFT_RIGHT))
            {
                cell = worksheet.Cells[coords.Item1, coords.Item2];
                if (!FormulaManager.IsEmptyCell(cell)) //if this is a data cell
                {
                    cell = worksheet.Cells[summaryRow, coords.Item2];
                    cell.SetCellValue(0, 0, 0.0);
                }
            }



            //Copy over all styles to the new row
            int lastColumn = worksheet.Dimension.End.Column;
            ExcelRange styleSource = worksheet.Cells[summaryRow - 1, 1, summaryRow - 1, lastColumn];
            ExcelRange destination = worksheet.Cells[summaryRow, 1, summaryRow, lastColumn];
            styleSource.CopyStyles(destination);


            headerCell.Style.Font.Bold = true;
        }




        /// <summary>
        /// Adds the new summary column to the end of the worksheet, gives it a header and copys over all the styles
        /// </summary>
        /// <param name="worksheet">the worksheet in need of formulas</param>
        /// <param name="headerRow">the row number of where the new column should have its header</param>
        private void AddSummaryColumn(ExcelWorksheet worksheet, int headerRow)
        {
            worksheet.Cells[headerRow, worksheet.Dimension.End.Column + 1].SetCellValue(0, 0, "Total:");


            //copy all formatting
            worksheet.Column(worksheet.Dimension.End.Column).AutoFit();

            int lastColumn = worksheet.Dimension.End.Column;
            int lastRow = worksheet.Dimension.End.Row;
            ExcelRange styleSource = worksheet.Cells[headerRow, lastColumn - 1, lastRow, lastColumn - 1];
            ExcelRange destination = worksheet.Cells[headerRow, lastColumn, lastRow, lastColumn];
            styleSource.CopyStyles(destination);



            //fill column with default values
            ExcelIterator iter = new ExcelIterator(worksheet, headerRow + 1, lastColumn);
            foreach(ExcelRange cell in iter.GetCells(ExcelIterator.SHIFT_DOWN))
            {
                cell.SetCellValue(0, 0, 0.0);
            }
        }




        public void SetDataCellDefenition(IsDataCell isDataCell)
        {
            this.dataCellDef = isDataCell;
        }
    }
}
