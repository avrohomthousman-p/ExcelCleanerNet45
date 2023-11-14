using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelCleanerNet45.GeneralCleaning
{

    /// <summary>
    /// An extentsion of the primary merge cleaner that ensures that all data cells (including the header) in a given 
    /// data column all end up in the same column after the unmerge.
    /// 
    /// Some reports (like ProfitAndLossBudget) have most of their data cells of a column merged over the same area, but
    /// then some cells (usually the column header) merged over a different area. This causes the column to not be aligned
    /// correctly after running the primary merge cleaner. The purpose of this class is to correct that issue after the
    /// unmerge, by moving cells into the data column that is nearest to their current location.
    /// </summary>
    class ReAlignMergeCells : PrimaryMergeCleaner
    {

        //used to store which column numbers the actual data columns are
        private HashSet<int> dataCols = null;



        public override void Unmerge(ExcelWorksheet worksheet)
        {
            FindTableBounds(worksheet);

            UnMergeMergedSections(worksheet);

            ReAlignWorksheet(worksheet); //this is where this object differs from the parent class

            ResizeCells(worksheet);

            DeleteColumns(worksheet);

            AdditionalCleanup(worksheet);
        }




        /// <summary>
        /// Executes the reAlignment of each data column in the worksheet
        /// </summary>
        /// <param name="worksheet">the worksheet currently being cleaned</param>
        private void ReAlignWorksheet(ExcelWorksheet worksheet)
        {
            dataCols = FindDataColumns(worksheet);




            //start from the first column with data cells
            int col = base.mergeRangesOfDataCells.Min(range => range.Item1);

            for(; col <= worksheet.Dimension.End.Column; col++)
            {
                ReAlignColumn(worksheet, col);
            }
        }



        //for debugging only
        private string GetRowLetter(int col)
        {
            string alphebet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

            col--;
            string result = "" + alphebet[col % 26];

            col /= 26;

            if(col > 0)
            {
                result = "A" + result;
            }

            return result;
        }




        /// <summary>
        /// Finds all the data columns in the worksheet
        /// </summary>
        /// <param name="worksheet">the worksheet we are currently cleaning</param>
        /// <returns>a Set with the column numbers for each data column in the worksheet</returns>
        private HashSet<int> FindDataColumns(ExcelWorksheet worksheet)
        {
            int rowsInWorksheet = worksheet.Dimension.End.Row - base.firstRowOfTable;
            HashSet<int> dataColumns = new HashSet<int>();


            for(int col = 1; col <= worksheet.Dimension.End.Column; col++)
            {
                int numDataCells = CountDataCellsInColumn(worksheet, col);

                if(HasManyDataCells(numDataCells, rowsInWorksheet))
                {
                    dataColumns.Add(col);
                }
            }


            return dataColumns;
        }



        /// <summary>
        /// Counts the number of non empty cells in the table (no major headers) that are inside the specified column
        /// </summary>
        /// <param name="worksheet">the worksheet currently being cleaned</param>
        /// <param name="column">the column to be counted</param>
        /// <returns>the number of non empty cells in the column</returns>
        private int CountDataCellsInColumn(ExcelWorksheet worksheet, int column)
        {
            ExcelIterator iter = new ExcelIterator(worksheet, base.firstRowOfTable, column);

            return iter.GetCells(ExcelIterator.SHIFT_DOWN).Count(cell => !IsEmptyCell(cell));
        }



        /// <summary>
        /// Checks if the column has a large number of data cells relative to the number of rows it has
        /// </summary>
        /// <param name="numCells">the number of data cells found in the column</param>
        /// <param name="rowsInWorksheet">the total number of cells found in the column</param>
        /// <returns>true if the column has a large number of data cells and false otherwise</returns>
        private bool HasManyDataCells(int numCells, int rowsInWorksheet)
        {
            return numCells >= (int)(.50 * rowsInWorksheet);
        }



        /// <summary>
        /// Moves all data found inside the specified column into the nearest data column. If the specified
        /// column is already a data column, this function will do nothing.
        /// </summary>
        /// <param name="worksheet">the worksheet currently being cleaned</param>
        /// <param name="column">the column we are cleaning</param>
        private void ReAlignColumn(ExcelWorksheet worksheet, int column)
        {
            if (dataCols.Contains(column)) //if this column is already a data column, we want to keep the data in it
            {
                return;
            }

            
            //Get the nearest and second to nearest data columns
            var nearestDataCol = GetNextNearestDataColumn(worksheet, column);

            int nearest = nearestDataCol.First();
            int secondNearest = nearestDataCol.Skip(1).First();



            ExcelRange sourceCell;
            ExcelRange destCell;
            for(int row = base.firstRowOfTable; row <= worksheet.Dimension.End.Row; row++)
            {
                sourceCell = worksheet.Cells[row, column];

                if (!IsEmptyCell(sourceCell))
                {


                    destCell = GetDestinationCell(worksheet, row, nearest, secondNearest);
                    if(destCell != null)
                    {
                        MoveCellToDataColumn(sourceCell, destCell);
                    }
                    else
                    {
                        Console.WriteLine($"Cell {sourceCell.Address} cannot be moved as the destination cell isnt empty");
                    }

                    
                }
            }
        }




        /// <summary>
        /// Finds the next data column that is closest to the specified column
        /// </summary>
        /// <param name="worksheet">the worksheet being cleaned</param>
        /// <param name="originalCol">the origin column that needs its nearest data column</param>
        /// <returns>the next closest column to the origin</returns>
        private IEnumerable<int> GetNextNearestDataColumn(ExcelWorksheet worksheet, int originalCol)
        {
            int leftSide = originalCol;
            int rightSide = originalCol;

            while (rightSide <= worksheet.Dimension.End.Column || leftSide >= 1)
            {

                if(dataCols.Contains(rightSide))
                {
                    yield return rightSide;
                }
                else if(dataCols.Contains(leftSide))
                {
                    yield return leftSide;
                }


                rightSide++;
                leftSide--;
            }
        }




        /// <summary>
        /// Gets the destination cell that a data cell should be moved to.
        /// </summary>
        /// <param name="worksheet">the worksheet currently being cleaned</param>
        /// <param name="row">the row the source cell is in</param>
        /// <param name="dataCol">the data column we want to move the cell to</param>
        /// <param name="backupDataCol">the backup column we should move the cell to if the other data column isnt availible</param>
        /// <returns>the cell the data should be moved to, or null if that cell isnt availible</returns>
        private ExcelRange GetDestinationCell(ExcelWorksheet worksheet, int row, int dataCol, int backupDataCol)
        {
            ExcelRange destCell = worksheet.Cells[row, dataCol];

            if (IsEmptyCell(destCell))
            {
                return destCell;
            }
            else
            {
                destCell = worksheet.Cells[row, backupDataCol];

                if (IsEmptyCell(destCell))
                {
                    return destCell;
                }
                else
                {
                    return null;
                }
            }
        }




        /// <summary>
        /// Moves the contents of the source cell to the data cell (if its empty) and ensures all styles are maintained
        /// </summary>
        /// <param name="source">the cell whose contents are to be moved</param>
        /// <param name="dest">the cell the data should be placed in</param>
        private void MoveCellToDataColumn(ExcelRange source, ExcelRange dest)
        {
            if (!IsEmptyCell(dest))
            {
                Console.WriteLine($"Data was supposed to be copied from {source.Address} to {dest.Address} but the cell was not empty");
                return;
            }


            dest.Value = source.Value;

            source.Value = null;

            base.CopyCellStyles(source, dest);
        }
    }
}
