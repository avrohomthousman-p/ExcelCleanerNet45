using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace ExcelCleanerNet45
{

    /// <summary>
    /// Implementation of the IMergeCleaner interface that seeks to maintain the exact same formatting after doing an
    /// unmerge. This is the primary implementation of this interface and should be the first tool used to clean merges.
    /// This implementation considers the table to start at the first row that contains at lease 3 non-empty cells.
    /// Cells are considered to be data cells if the column they are in contains text in the cell at the top of the table.
    /// All major and minor headers are not resized, but get set to not wrap their text, so excel can display it outside
    /// its cell boundry. Data cells are resized such that their size after the unmerge matches their original size before
    /// the unmerge.
    /// </summary>
    internal class PrimaryMergeCleaner : AbstractMergeCleaner
    {

        //number of non empty cells to be the first data row
        private static readonly int NUM_FULL_COLUMNS_REQUIRED = 3;


        protected int firstRowOfTable = -1;


        //This stores the start and end columns of the merges that the header of each data column are in.
        //Any cell inside the table, whose merge spans the same cells as any entry in this Set, is 
        //considered a data cell.
        //
        //Example: if the worksheet has a data column with all it's data cells merged over columns 3-6,
        //then this hashset will have an entry of (3,6)
        protected HashSet<Tuple<int, int>> mergeRangesOfDataCells;


        private Dictionary<int, double> originalColumnWidths = new Dictionary<int, double>();





        /// <inheritdoc/>
        protected override void FindTableBounds(ExcelWorksheet worksheet)
        {
            /*
             * We want the first data row to be the first row that has at least 3 non empty cells
             * in it, and also has a border at its bottom. If there are no rows that have a border at the bottom
             * instead just take the first row with 3 non-empty cells in it.
             */

            bool foundRowWith3Values = false;

            for (int row = 1; row <= worksheet.Dimension.End.Row; row++)
            {
                if (IsDataRow(worksheet, row))
                {

                    if (RowHasBottomBorder(worksheet, row))
                    {
                        firstRowOfTable = row;
                        break;
                    }
                    else if (!foundRowWith3Values)
                    {
                        firstRowOfTable = row;
                        foundRowWith3Values = true;
                    }

                }

            }


            if(firstRowOfTable == -1)
            {
                //if no data row is found there is something wrong
                throw new System.IO.InvalidDataException("Could not find data table in excel report.");
            }
            else
            {
                Console.WriteLine("First data row is row " + firstRowOfTable);

                TrackDataColumns(worksheet);
            }
        }



        /// <summary>
        /// Checks if the row has a bottom border
        /// </summary>
        /// <param name="worksheet">the worksheet currently being cleaned</param>
        /// <param name="row">the row whose border is being checked</param>
        /// <returns>true if the specifed row has a border on its bottom side, and false otherwise</returns>
        private bool RowHasBottomBorder(ExcelWorksheet worksheet, int row)
        {
            var cell = worksheet.Cells[row, 2];

            var cellBorder = cell.Style.Border.Bottom.Style;

            return !cellBorder.Equals(ExcelBorderStyle.None);
        }




        /// <summary>
        /// Checks if the specified row is a data row.
        /// </summary>
        /// <param name="worksheet">the worksheet where the row in question can be found</param>
        /// <param name="row">the row to be checked</param>
        /// <returns>true if the specified row is a data row and false otherwise</returns>
        private bool IsDataRow(ExcelWorksheet worksheet, int row)
        {

            int fullColumnsFound = 0;


            for (int col = 1; col <= worksheet.Dimension.Columns; col++)
            {
                if (!IsEmptyCell(worksheet.Cells[row, col]))
                {
                    fullColumnsFound++;
                    if (fullColumnsFound == NUM_FULL_COLUMNS_REQUIRED)
                    {
                        return true;
                    }
                }
            }

            return false;
        }




        /// <summary>
        /// Populates the local variable isDataColumn with true's for each column that is a data column
        /// </summary>
        /// <param name="worksheet">the worksheet we are currently cleaning</param>
        private void TrackDataColumns(ExcelWorksheet worksheet)
        {

            mergeRangesOfDataCells = new HashSet<Tuple<int, int>>();

            for(int col = 1; col <= worksheet.Dimension.Columns; col++)
            {
                ExcelRange fullMerge = GetMergeCellByPosition(worksheet, firstRowOfTable, col);


                if(fullMerge == null) //If the current cell is not a merge cell
                {
                    continue;
                }
                else if (IsEmptyCell(fullMerge))
                {
                    continue; //skip this cell
                }


                mergeRangesOfDataCells.Add(new Tuple<int, int>(fullMerge.Start.Column, fullMerge.End.Column));
                col = fullMerge.End.Column;
            }

        }




        /// <inheritdoc/>
        protected override void UnMergeMergedSections(ExcelWorksheet worksheet)
        {
            Console.WriteLine("unmerging worksheet " + worksheet.Index);
            RecordOriginalColumnWidths(worksheet);


            ExcelWorksheet.MergeCellsCollection mergedCells = worksheet.MergedCells;


            for (int i = mergedCells.Count - 1; i >= 0; i--)
            {
                var merged = mergedCells[i];


                //sometimes a change to one part of the worksheet causes a merge cell to stop
                //existing. The corrisponding entry in the merge collection to becomes null.
                if (merged == null)
                {
                    continue;
                }

                UnMergeCells(worksheet, merged.ToString());
            }

        }



        /// <summary>
        /// Records the starting widths of all data columns that are merged in a dictionary for later use in resizing those columns
        /// </summary>
        /// <param name="worksheet">the worksheet currently being cleaned</param>
        private void RecordOriginalColumnWidths(ExcelWorksheet worksheet)
        {

            for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
            {

                ExcelRange currentCell = worksheet.Cells[firstRowOfTable, col];


                if (!currentCell.Merge || IsEmptyCell(currentCell))
                {
                    continue;
                }


                currentCell = GetMergeCellByPosition(worksheet, firstRowOfTable, col);

                double width = GetWidthOfMergeCell(worksheet, currentCell);

                originalColumnWidths.Add(col, width);

                col += (CountMergeCellLength(currentCell) - 1);
            }
        }



        /// <summary>
        /// Calculates the width of a merge cell.
        /// </summary>
        /// <param name="worksheet">the worksheet currently being cleaned</param>
        /// <param name="currentCells">the merge cell being mesured</param>
        /// <returns>the width of the specified merged cell</returns>
        private double GetWidthOfMergeCell(ExcelWorksheet worksheet, ExcelRange currentCells)
        {
            double width = 0;

            for (int col = currentCells.Start.Column; col <= currentCells.End.Column; col++)
            {
                width += worksheet.Column(col).Width; //alt:  currentCell.EntireColumn.Width;  
            }


            return width;
        }



        /// <summary>
        /// Counts the number of cells in a merge cell. This method only counts the number of cells
        /// in the first row of the merge.
        /// </summary>
        /// <param name="mergeCell">the address of the full merge cell</param>
        /// <returns>the number of cells in the merge</returns>
        private int CountMergeCellLength(ExcelRange mergeCell)
        {
            return mergeCell.End.Column - mergeCell.Start.Column + 1;
        }




        /// <summary>
        /// Unmerges the specified segment of merged cells.
        /// </summary>
        /// <param name="worksheet">the worksheet we are currently cleaning</param>
        /// <param name="cellAddress">the address of the ENTIRE merged section (eg A18:F24)</param>
        /// <returns>true if the specified cell was unmerged, and false otherwise</returns>
        private bool UnMergeCells(ExcelWorksheet worksheet, string cellAddress)
        {

            ExcelRange currentCells = worksheet.Cells[cellAddress];


            //record the style we had before any changes were made
            ExcelStyle originalStyle = currentCells.Style;




            MergeType mergeType = GetCellMergeType(currentCells);


            switch (mergeType)
            {
                case MergeType.NOT_A_MERGE:
                    return false;

                case MergeType.EMPTY:
                    break;

                case MergeType.MAIN_HEADER:
                    currentCells.Style.WrapText = false;
                    ConvertContentsToText(currentCells); //Ensure that dates are displayed correctly
                    Console.WriteLine("major header at " + currentCells.Address);
                    break;

                case MergeType.MINOR_HEADER:
                    currentCells.Style.WrapText = false;
                    break;

                case MergeType.DATA:
                    break;
            }


            //turning off custom hieght before unmerge will allow row to resize itself to fit everything
            worksheet.Row(currentCells.Start.Row).CustomHeight = false;




            //unmerge range
            currentCells.Merge = false;




            //restore the original style
            SetCellStyles(currentCells, originalStyle);
            


            //If there is more than one line of text in a header, it should be split into
            //multiple seperate headers. Also, if the header is not aligned left, the text 
            //should not be moved to the leftmost side of the formally merged range.
            if(mergeType == MergeType.MAIN_HEADER)
            {
                SplitHeaderIntoMultipleRows(worksheet, currentCells);
                MoveMajorHeadersToMatchFormatting(worksheet, currentCells);

            }
            //If there is a minor header that is aligned right, it usually should
            //stay on the right side of the now unmerged range
            else if(mergeType == MergeType.MINOR_HEADER)
            {
                MoveMinorHeaderToMatchFormatting(worksheet, currentCells);
            }



            return true;

        }




        /// <summary>
        /// Gets the type of merge that is found in the specified cell
        /// </summary>
        /// <param name="cell">the cell whose merge type is being checked</param>
        /// <returns>the MergeType object that corrisponds to the type of merge cell we are given</returns>
        private MergeType GetCellMergeType(ExcelRange cell)
        {
            if (cell.Merge == false)
            {
                return MergeType.NOT_A_MERGE;
            }

            if (IsEmptyCell(cell))
            {
                return MergeType.EMPTY;
            }

            if (IsMajorHeader(cell))
            {
                return MergeType.MAIN_HEADER;
            }

            if (IsMinorHeader(cell))
            {
                return MergeType.MINOR_HEADER;
            }

            if (IsDataCell(cell))
            {
                return MergeType.DATA;
            }


            throw new ArgumentException("Cannot determan merge type of specified cell");
        }



        /// <summary>
        /// Moves the text of the major header to the center of the formally merged range if it is aligned center,
        /// and to the right if it is aligned right.
        /// </summary>
        /// <param name="worksheet">the worksheet currently being cleaned</param>
        /// <param name="cellRange">the range that was just unmerged</param>
        private void MoveMajorHeadersToMatchFormatting(ExcelWorksheet worksheet, ExcelRange cellRange)
        {
            if (cellRange.Style.HorizontalAlignment.Equals(ExcelHorizontalAlignment.Left))
            {
                return;
            }


            int row = cellRange.Start.Row;
            int startCol = cellRange.Start.Column;
            int endCol = cellRange.End.Column;

            ExcelRange source = worksheet.Cells[row, startCol];
            ExcelRange destination;



            if (cellRange.Style.HorizontalAlignment.Equals(ExcelHorizontalAlignment.Center))
            {
                int middleColumn = startCol + ((endCol - startCol) / 2);
                destination = worksheet.Cells[row, middleColumn];
            }
            else if (cellRange.Style.HorizontalAlignment.Equals(ExcelHorizontalAlignment.Right))
            {
                destination = worksheet.Cells[row, endCol];
            }
            else
            {
                return;
            }



            if(source.Address == destination.Address)
            {
                return;
            }

            source.Copy(destination);
            source.CopyStyles(destination);

            source.Value = null;
        }



        /// <summary>
        /// Moves minor headers that were aligned right, to the rightmost cell in the specified cell range, 
        /// and minor headers aligned center to the center of the range.
        /// </summary>
        /// <param name="worksheet">the worksheet currently being cleaned</param>
        /// <param name="cellRange">the original range that held the minor header before the unmerge</param>
        private void MoveMinorHeaderToMatchFormatting(ExcelWorksheet worksheet, ExcelRange cellRange)
        {
            if (cellRange.Style.HorizontalAlignment.Equals(ExcelHorizontalAlignment.Left))
            {
                return;
            }


            int row = cellRange.Start.Row;
            int startCol = cellRange.Start.Column;
            int endCol = cellRange.End.Column;

            ExcelRange source = worksheet.Cells[row, startCol];
            ExcelRange destination;



            //if it is actually a data cell, its just considered a minor header because its in the wrong row
            //then it gets moved to the nearest data column
            if (source.Text.StartsWith("$") || (source.Text.StartsWith("($") && source.Text.EndsWith(")")))
            {
                int endOfDataColumn = GetNearestDataColumn(startCol, endCol).Item1;
                destination = worksheet.Cells[row, endOfDataColumn];

                //if we cant move the data there becuase its in use
                if (!SafeToCopyTextTo(destination)) 
                {
                    destination = worksheet.Cells[row, endCol];
                }

            }
            else if(cellRange.Style.HorizontalAlignment.Equals(ExcelHorizontalAlignment.Right))
            {
                destination = worksheet.Cells[row, endCol];
            }
            else if (cellRange.Style.HorizontalAlignment.Equals(ExcelHorizontalAlignment.Center))
            {
                int middleColumn = startCol + ((endCol - startCol) / 2);
                destination = worksheet.Cells[row, middleColumn];
            }
            else
            {
                return;
            }



            //Copy the data
            if(source.Address != destination.Address)
            {
                source.Copy(destination);
                source.CopyStyles(destination);

                source.Value = null;
            }
        }



        /// <summary>
        /// Checks if it is safe to copy new text into the specified cell
        /// </summary>
        /// <param name="destination">the cell wherethe new text needs to go</param>
        /// <returns>true if it is safe to copy the text, and false otherwise</returns>
        private bool SafeToCopyTextTo(ExcelRange destination)
        {
            return !destination.Merge && IsEmptyCell(destination);
        }



        /// <summary>
        /// Given the span of a merge cell, gets the span of the data column that most closely matches that merge cell,
        /// or the span of the merge cell itself if no matches are found.
        /// </summary>
        /// <param name="startCol">the starting column of the merge cell</param>
        /// <param name="endCol">the ending column of the merge cell</param>
        /// <returns>
        /// a tuple with the starting and ending columns of the data column most similar to the specified merge span, or
        /// the span of the merge itself if no appropriate data column is found.
        /// </returns>
        private Tuple<int, int> GetNearestDataColumn(int startCol, int endCol)
        {
            foreach(Tuple<int, int> column in mergeRangesOfDataCells)
            {
                if (MergeMatchesDataColumn(startCol, endCol, column))
                {
                    return column;
                }
            }


            return new Tuple<int, int>(startCol, endCol);
        }



        /// <summary>
        /// Checks if the specified merge span closely matches the specified data column span
        /// </summary>
        /// <param name="mergeStart">the column number that the merge starts in</param>
        /// <param name="mergeEnd">the column number that the merge ends in</param>
        /// <param name="dataColumn">a tuple containing the start and end column of the data column span</param>
        /// <returns>true if the data column and merge cell span mostly the same area</returns>
        private bool MergeMatchesDataColumn(int mergeStart, int mergeEnd, Tuple<int, int> dataColumn)
        {
            int numOverlappingCols;

            //now count how many columns are shared in both spans (overlap)

            //if the merge cell is a bit to the right of the data column
            if(mergeStart >= dataColumn.Item1 &&  mergeEnd >= dataColumn.Item2) 
            {
                numOverlappingCols = dataColumn.Item2 - mergeStart;
            }
            //if the merge cell is a bit to the left of the data column
            else if (mergeStart <= dataColumn.Item1 && mergeEnd <= dataColumn.Item2)
            {
                numOverlappingCols =  mergeEnd - dataColumn.Item1;
            }
            //if there is no overlap
            else if (dataColumn.Item2 < mergeStart || dataColumn.Item1 > mergeEnd) 
            {
                numOverlappingCols = 0;
            }
            //if there is complete overlap except that one is bigger than the other
            else
            {
                numOverlappingCols = dataColumn.Item2 - dataColumn.Item1;
            }


            //if at least 2/3 of the cells overlap return true
            return ((double)numOverlappingCols) / (dataColumn.Item2 - dataColumn.Item1) >= 0.66;
        }



        /// <inheritdoc/>
        protected override void ResizeCells(ExcelWorksheet worksheet)
        {
            //resize all columns
            foreach (KeyValuePair<int, double> entry in originalColumnWidths)
            {
                worksheet.Column(entry.Key).Width = entry.Value;
            }
        }




        /// <inheritdoc/>
        protected override void DeleteColumns(ExcelWorksheet worksheet)
        {

            //We don't want to delete any columns before the first data column becuase that might 
            //mess up the whitespace around minor headers
            int firstDataColumn = FindFirstDataColumn(worksheet);


            for (int col = worksheet.Dimension.Columns; col > firstDataColumn; col--)
            {
                if (SafeToDeleteColumn(worksheet, col))
                {
                    PrepareColumnForDeletion(worksheet, col);
                    worksheet.DeleteColumn(col);

                    Console.WriteLine("Column " + col + " is being deleted");
                }
            }
        }



        /// <summary>
        /// Finds the column number of the first column that is considered a data column
        /// </summary>
        /// <param name="worksheet">the worksheet currently being cleaned</param>
        /// <returns>the column number of the first data column</returns>
        private int FindFirstDataColumn(ExcelWorksheet worksheet)
        {
            for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
            {
                ExcelRange cell = worksheet.Cells[firstRowOfTable, col];
                if (!IsEmptyCell(cell))
                {
                    return col;
                }
            }


            return 1; //Default: first column
        }



        /// <summary>
        /// Checks if a column is safe to delete becuase it is empty other than possibly having major headers in it.
        /// </summary>
        /// <param name="worksheet">the worksheet where the column can be found</param>
        /// <param name="col">the column being checked</param>
        /// <returns>true if it is safe to delete the column and false if deleting it would result in data loss</returns>
        private bool SafeToDeleteColumn(ExcelWorksheet worksheet, int col)
        {
            for (int row = firstRowOfTable; row <= worksheet.Dimension.End.Row; row++)
            {
                if (!IsEmptyCell(worksheet.Cells[row, col]))
                {
                    return false;
                }
            }


            return true;
        }



        /// <summary>
        /// Moves all major headers in the specified column to the column adjacent on the left, or right if we are on the  
        /// first column.
        /// </summary>
        /// <param name="worksheet">the worksheet the column could be found in</param>
        /// <param name="col">the column number we are preparing to delete</param>
        private void PrepareColumnForDeletion(ExcelWorksheet worksheet, int col)
        {
            for (int row = 1; row < firstRowOfTable; row++)
            {
                if (!IsEmptyCell(worksheet.Cells[row, col]))
                {

                    int destinationColumn;
                    if (col == 1) //we are on the first column
                    {
                        destinationColumn = 2; //move header right
                    }
                    else
                    {
                        destinationColumn = col - 1; //move header left
                    }

                    ExcelRange originCell = worksheet.Cells[row, col];
                    ExcelRange destinationCell = worksheet.Cells[row, destinationColumn];


                    //copy all styles and formatting
                    originCell.CopyStyles(destinationCell);

                    //Move the text to the destination cell (store it as a string to avoid excel display issues with dates)
                    destinationCell.SetCellValue(0, 0, originCell.Text);


                    originCell.Value = null;
                }
            }
        }



        /// <inheritdoc/>
        protected override void AdditionalCleanup(ExcelWorksheet worksheet)
        {
            base.AdditionalCleanup(worksheet);

            MoveMajorHeadersLeft(worksheet, firstRowOfTable);
        }




        /// <inheritdoc/>
        protected override bool IsDataCell(ExcelRange cell)
        {

            int start = cell.Start.Column;
            int end = cell.End.Column;
            return mergeRangesOfDataCells.Contains(new Tuple<int, int>(start, end));

            //return isDataColumn[cell.Start.Column - 1];

        }



        /// <inheritdoc/>
        protected override bool IsInsideTable(ExcelRange cell)
        {

            return cell.Start.Row >= firstRowOfTable;

        }



        /// <inheritdoc/>
        protected override bool IsMajorHeader(ExcelRange cell)
        {
            return !IsEmptyCell(cell) && cell.Start.Row < firstRowOfTable;
        }



        /// <inheritdoc/>
        protected override bool IsMinorHeader(ExcelRange cells)
        {
            if (IsEmptyCell(cells) || !IsInsideTable(cells))
            {
                return false;
            }



            //return !isDataColumn[cells.Start.Column - 1];
            return !IsDataCell(cells);
        }
    }
}

