using OfficeOpenXml;
using System;
using System.Linq;

namespace ExcelCleanerNet45.GeneralCleaning
{
    /// <summary>
    /// A version of the BackupMergeCleaner that adds additional cleanup specific to the 
    /// ProfitAndLossExtendedVarianece report.
    /// </summary>
    internal class ExtendedVarianceCleaner : BackupMergeCleaner
    {
        /// <inheritdoc/>
        protected override void AdditionalCleanup(ExcelWorksheet worksheet)
        {
            base.AdditionalCleanup(worksheet);

            MoveFinalSummaryCells(worksheet);
        }



        /// <summary>
        /// Moves the summary cells at the bottom of the ProfitAndLossExtendedVarianece report so that
        /// they align with the proper data columns.
        /// </summary>
        /// <param name="worksheet">the worksheet being cleaned</param>
        protected void MoveFinalSummaryCells(ExcelWorksheet worksheet)
        {
            int referenceRow =  base.topTableRow; //we want the bottom row to align with the reference row
            int mainRow = FindLastNonEmptyRow(worksheet);
            Console.WriteLine($"main row = {mainRow}, reference row = {referenceRow}");

            ExcelRange referenceCell, actualCell;

            var referenceIter = new ExcelIterator(worksheet, referenceRow, 1).GetCells(ExcelIterator.SHIFT_RIGHT);
            var mainIter = new ExcelIterator(worksheet, mainRow, 1).GetCells(ExcelIterator.SHIFT_RIGHT);

            bool keepGoing = true;



            while(keepGoing)
            {
                referenceCell = referenceIter.SkipWhile(cell => base.IsEmptyCell(cell) || !base.IsDataCell(cell)).FirstOrDefault();
                actualCell = mainIter.SkipWhile(cell => base.IsEmptyCell(cell) || !base.IsDataCell(cell)).FirstOrDefault();

                if(referenceCell == null || actualCell == null)
                {
                    keepGoing = false;
                }
                else
                {
                    Console.WriteLine($"reference cell = {referenceCell.Address} main cell = {actualCell}");
                    MoveCellIfNecessary(worksheet, referenceCell, actualCell);

                    //This code seems to fix some strange bug in the IEnumerable functions
                    referenceIter = referenceIter.Skip(1);
                    mainIter = mainIter.Skip(1);
                }
            }
        }




        /// <summary>
        /// Finds the last row in the worksheet that has data cells in it
        /// </summary>
        /// <param name="worksheet">the worksheet being cleaned</param>
        /// <returns>the row number of the last row containing data</returns>
        protected int FindLastNonEmptyRow(ExcelWorksheet worksheet)
        {
            ExcelIterator iter = new ExcelIterator(worksheet, worksheet.Dimension.End.Row, worksheet.Dimension.End.Column);
            return iter.FindAllCellsReverse().First(cell => base.IsDataCell(cell)).Start.Row;
        }




        /// <summary>
        /// Moves all data and formatting from the origin cell to a cell on the same row, but in the same column
        /// as the reference cell.
        /// </summary>
        /// <param name="worksheet">the worksheet currently being cleaned</param>
        /// <param name="referenceCell">the cell in the colum the data should be in</param>
        /// <param name="originCell">the actual location of the data</param>
        protected void MoveCellIfNecessary(ExcelWorksheet worksheet, ExcelRange referenceCell, ExcelRange originCell)
        {
            //if these cells are already aligned
            if (referenceCell.Start.Column == originCell.Start.Column)
            {
                return;
            }


            //move origin cell into the same column as the reference cell if the reference cell is empty
            ExcelRange destinationCell = worksheet.Cells[originCell.Start.Row, referenceCell.Start.Column];

            if (base.IsEmptyCell(destinationCell))
            {
                originCell.Copy(destinationCell);
                originCell.CopyStyles(destinationCell);
                originCell.Value = null;
            }
        }
    }
}
