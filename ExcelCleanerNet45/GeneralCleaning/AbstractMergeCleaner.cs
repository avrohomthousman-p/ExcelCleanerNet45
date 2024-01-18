using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;

namespace ExcelCleanerNet45
{

    /// <summary>
    /// Defines the basic layout for cleaning merges in an excel file
    /// </summary>
    internal abstract class AbstractMergeCleaner : IMergeCleaner
    {

        //Used to track all the jobs that must be done in the AdditionalCleanup method
        private List<Action<ExcelWorksheet>> Cleanups = new List<Action<ExcelWorksheet>>();


        protected bool moveMajorHeaders = true;

        /// <summary>
        /// If set to true, this object will move all major headers near the left side of the worksheet into column 1
        /// </summary>
        public bool MoveMajorHeaders
        {
            get { return moveMajorHeaders; }
            set { moveMajorHeaders = value; }
        }


        //determans how many columns get moved left.
        //e.g. if set to 3, only the first 3 columns will get moved left
        protected int columnsToMove = 3;



        /// <summary>
        /// Represents the different ways that borders can be removed from the top section of the report.
        /// </summary>
        internal enum BorderRemovalType { NONE, ALL, ONLY_EMPTY_CELLS }


        protected BorderRemovalType borderRemovalSelection = BorderRemovalType.ONLY_EMPTY_CELLS;


        /// <summary>
        /// Controls how border lines are removed from the major header section of the report.
        /// </summary>
        public BorderRemovalType BorderRemoval
        {
            get { return borderRemovalSelection; }
            set { borderRemovalSelection = value; }
        }




        /// <summary>
        /// Sets the number of columns whose major headers will be moved to the start of the page.
        /// For example, this function is called with a 3 passed in, only major headers found in the first 
        /// 3 columns will be moved to column 1.
        /// 
        /// Note: if the MoveMajorHeaders property is turned off, calling this function will not have any effect
        /// on the program output.
        /// </summary>
        /// <param name="columns">the number of columns whose major headers should be moved</param>
        public virtual void SetNumColumnsToMove(int columns)
        {
            columnsToMove = columns;
        }





        public virtual void Unmerge(ExcelWorksheet worksheet)
        {
            FindTableBounds(worksheet);

            UnMergeMergedSections(worksheet);

            ResizeCells(worksheet);

            DeleteColumns(worksheet);

            AdditionalCleanup(worksheet);
        }




        /// <summary>
        /// Finds the first row that is considered part of the table in the specified worksheet and saves the 
        /// row number to a local variable for later use
        /// </summary>
        /// <param name="worksheet">the worksheet we are working on</param>
        /// <exception cref="Exception">if the first row of the table couldnt be found</exception>
        protected abstract void FindTableBounds(ExcelWorksheet worksheet);



        /// <summary>
        /// Unmerges all the merged sections in the worksheet.
        /// </summary>
        /// <param name="worksheet">the worksheet we are currently cleaning</param>
        protected abstract void UnMergeMergedSections(ExcelWorksheet worksheet);



        /// <summary>
        /// Resizes all columns and rows that need a resize
        /// </summary>
        /// <param name="worksheet">the worksheet currently being cleaned</param>
        protected abstract void ResizeCells(ExcelWorksheet worksheet);



        /// <summary>
        /// Deletes all empty/unwanted columns in the table
        /// </summary>
        /// <param name="worksheet">the worksheet currently being cleaned</param>
        protected abstract void DeleteColumns(ExcelWorksheet worksheet);



        /// <summary>
        /// Adds a cleanup job to the list of cleanup actions that must be done
        /// </summary>
        /// <param name="job">the job that must be executed during the initial cleanup</param>
        internal void AddCleanupJob(Action<ExcelWorksheet> job)
        {
            this.Cleanups.Add(job);
        }



        /// <summary>
        /// Does all additional cleanup that is needed for the specified report
        /// </summary>
        /// <param name="worksheet">the worksheet being cleaned</param>
        /// <param name="reportName">the report being cleaned</param>
        protected virtual void AdditionalCleanup(ExcelWorksheet worksheet)
        {
            //execute any additional cleaning jobs that were added by the AddCleanupJob() function
            foreach (Action<ExcelWorksheet> job in Cleanups)
            {
                try
                {
                    job.Invoke(worksheet);
                }
                catch(Exception e)
                {
                    Console.WriteLine("Error while doing custom cleanup job. Message: ");
                    Console.WriteLine(e.Message);
                }
                
            }
        }




        /* Some Abstract Utility Methods */


        /// <summary>
        /// Checks if the cell at the specified coordinates is a data cell
        /// </summary>
        /// <param name="cell">the cell being checked</param>
        /// <returns>true if the cell is a data cell and false otherwise</returns>
        protected abstract bool IsDataCell(ExcelRange cell);



        /// <summary>
        /// Checks if the specified cell is inside the table in the worksheet, and not a header 
        /// above the table
        /// </summary>
        /// <param name="cell">the cell whose location is being checked</param>
        /// <returns>true if the specified cell is inside a table and false otherwise</returns>
        protected abstract bool IsInsideTable(ExcelRange cell);



        /// <summary>
        /// Checks if the specified cell is a major header.
        /// 
        /// A major header is defined as a header that is above the table.
        /// </summary>
        /// <param name="cell">the cell being checked</param>
        /// <returns>true if the specified cell contains a major header, and false otherwise</returns>
        protected abstract bool IsMajorHeader(ExcelRange cell);



        /// <summary>
        /// Checks if the specified cell is considered a minor header.
        /// 
        /// A minor header is defined as a merge cell that contains non-data text and is inside the table.
        /// </summary>
        /// <param name="cells">the cells that we are checking</param>
        /// <returns>true if the specified cells are a minor header and false otherwise</returns>
        protected abstract bool IsMinorHeader(ExcelRange cells);





        /* Some Utility Methods With Implementations */


        /// <summary>
        /// Checks if a cell has no text
        /// </summary>
        /// <param name="currentCells">the cell that is being checked for text</param>
        /// <returns>true if there is no text in the cell, and false otherwise</returns>
        protected bool IsEmptyCell(ExcelRange currentCells)
        {
            return currentCells.Text == null || currentCells.Text.Length == 0;
        }



        /// <summary>
        /// Sets the PatternType, Color, Border, Font, and Horizontal Alingment of all the cells
        /// in the specifed range.
        /// </summary>
        /// <param name="currentCells">the cells whose style must be set</param>
        /// <param name="style">all the styles we should use</param>
        protected virtual void SetCellStyles(ExcelRange currentCells, ExcelStyle style)
        {


            //Ensure the color of the cells dont get changed
            currentCells.Style.Fill.PatternType = style.Fill.PatternType;
            if (currentCells.Style.Fill.PatternType != ExcelFillStyle.None)
            {
                currentCells.Style.Fill.PatternColor.SetColor(
                    GetColorFromARgb(style.Fill.PatternColor.LookupColor()));

                currentCells.Style.Fill.BackgroundColor.SetColor(
                    GetColorFromARgb(style.Fill.BackgroundColor.LookupColor()));
            }


            currentCells.Style.Border = style.Border;
            currentCells.Style.Font.Bold = style.Font.Bold;
            currentCells.Style.Font.Size = style.Font.Size;
            currentCells.Style.Font.Name = style.Font.Name;
            currentCells.Style.Font.Scheme = style.Font.Scheme;
            currentCells.Style.Font.Charset = style.Font.Charset;


            //Not sure why but it only sucsessfully sets these settings if these 2 lines are NOT executed
            //currentCells.Style.WrapText = style.WrapText;
            //currentCells.Style.HorizontalAlignment = style.HorizontalAlignment;
        }




        /// <summary>
        /// Copys styles from the source to destination cell, but does not copy color, borders, or cell content
        /// </summary>
        /// <param name="source">the cell whose style we should use</param>
        /// <param name="destination">the cell that should recieve the style</param>
        protected virtual void CopyCellStyles(ExcelRange source, ExcelRange destination)
        {
            destination.Style.HorizontalAlignment = source.Style.HorizontalAlignment;
            destination.Style.VerticalAlignment = source.Style.VerticalAlignment;

            destination.Style.Font.Bold = source.Style.Font.Bold;
            destination.Style.Font.Size = source.Style.Font.Size;
            destination.Style.Font.Name = source.Style.Font.Name;
            destination.Style.Font.Scheme = source.Style.Font.Scheme;
            destination.Style.Font.Charset = source.Style.Font.Charset;

            destination.Style.WrapText = source.Style.WrapText;
        }



        /// <summary>
        /// Generates A Color Object from an ARGB string.
        /// </summary>
        /// <param name="argb">the argb code of the color needed</param>
        /// <returns>an instance of System.Drawing.Color that matches the specified argb code</returns>
        protected virtual System.Drawing.Color GetColorFromARgb(String argb)
        {
            if (argb.StartsWith("#"))
            {
                argb = argb.Substring(1);
            }
            else if (argb.StartsWith("0x"))
            {
                argb = argb.Substring(2);
            }


            System.Drawing.Color color = System.Drawing.Color.FromArgb(
                int.Parse(argb, System.Globalization.NumberStyles.HexNumber));


            return color;
        }



        /// <summary>
        /// Ensures that the data in the specified cell is stored as text, not as a number. This is usefull
        /// if you need to ensure that a date in the report header does not get displayed as hashtags if the 
        /// column is too small.
        /// </summary>
        /// <param name="cell">the cell whose data must be converted to text</param>
        protected virtual void ConvertContentsToText(ExcelRange cell)
        {
            cell.SetCellValue(0, 0, cell.Text);
        }



        /// <summary>
        /// Splits a header cell with more than one line of text, into multiple rows,
        /// one for each line of text. Note: this operation should be done AFTER unmerging the cell.
        /// </summary>
        /// <param name="worksheet">the worksheet currently being cleaned</param>
        /// <param name="cells">the header cell containing multi-line text</param>
        /// <returns>the number of new rows that were inserted</returns>
        protected virtual int SplitHeaderIntoMultipleRows(ExcelWorksheet worksheet, ExcelRange cells)
        {
            if (!cells.Text.Contains("\n"))
            {
                return 0;
            }


            string[] linesOfText = cells.Text.Split('\n');

            int numNewRows = linesOfText.Length - 1;
            int startRow = cells.Start.Row;
            int endRow = startRow + numNewRows;

            //add empty rows to fit the text on different lines
            worksheet.InsertRow(startRow + 1, numNewRows);


            for (int rowNum = startRow; rowNum <= endRow; rowNum++)
            {
                var currentCell = worksheet.Cells[rowNum, cells.Start.Column];
                int indexOfText = rowNum - startRow;
                currentCell.SetCellValue(0, 0, linesOfText[indexOfText]);

                CopyCellStyles(cells, currentCell);
            }


            //if the original cell had a bottom border, we want to move that bottom border to the last new row 

            worksheet.Cells[endRow, cells.Start.Column].Style.Border.Bottom.Style = cells.Style.Border.Bottom.Style;
            cells.Style.Border.Bottom.Style = ExcelBorderStyle.None; //remove the bottom border of the top cell


            return numNewRows;
        }



        /// <summary>
        /// Ensures that any major header that ends up in the first few columns, gets moved to column 1 (if possible).
        /// The number of columns whose headers will be moved can be adjusted by calling the SetNumColumnsToMove function.
        /// </summary>
        /// <param name="worksheet">the worksheet that is being cleaned</param>
        /// <param name="firstDataRow">the row that marks the beginning of the data section. All cells above it are major headers</param>
        protected virtual void MoveMajorHeadersLeft(ExcelWorksheet worksheet, int firstDataRow)
        {
            if (!moveMajorHeaders)
            {
                return;
            }


            int lastColumnBeingMoved = Math.Min(this.columnsToMove, worksheet.Dimension.End.Column);

            for (int col = 2; col <= lastColumnBeingMoved; col++)
            {
                for(int row = 1; row < firstDataRow; row++)
                {
                    MoveHeaderIfNeeded(worksheet, row, col);
                }
            }
        }




        /// <summary>
        /// Moves the header found at the specified coordinates to the first column of the worksheet if possible 
        /// </summary>
        /// <param name="worksheet">the worksheet being cleaned</param>
        /// <param name="row">the row number of the header cell</param>
        /// <param name="col">the column number of the header cell</param>
        /// <returns>true if the header was moved sucsessfully, and false otherwise</returns>
        protected virtual bool MoveHeaderIfNeeded(ExcelWorksheet worksheet, int row, int col)
        {
            ExcelRange source = worksheet.Cells[row, col];
            ExcelRange dest = worksheet.Cells[row, 1];

            if (!IsEmptyCell(dest) || IsEmptyCell(source))
            {
                return false;
            }


            source.CopyStyles(dest);
            dest.Value = source.Value;
            source.Value = null;

            //headers on the left side of the page must be aligned left or they display off the screen
            dest.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

            dest.Style.WrapText = false;

            return true;
        }




        /// <summary>
        /// Some reports have major headers that take up only one cell (they are not merge cells). Those
        /// headers do not get their settings adjusted in the unmerge function, so they might need
        /// settings adjustment.
        /// </summary>
        /// <param name="worksheet">the worksheet being cleaned</param>
        /// <param name="firstDataRow">the row above which all non empty cells are considered major headers</param>
        protected virtual void CleanNonMergedMajorHeaders(ExcelWorksheet worksheet, int firstDataRow)
        {
            ExcelRange cell;

            for(int row = 1; row < firstDataRow; row++)
            {
                for(int col = 1; col <= worksheet.Dimension.End.Column; col++)
                {
                    cell = worksheet.Cells[row, col];
                    if (!IsEmptyCell(cell))
                    {
                        cell.Style.WrapText = false;
                        ConvertContentsToText(cell); //Ensure that dates are displayed correctly
                    }
                }
            }
        }




        /// <summary>
        /// Removes all cell borders above the data table that are not the decorating a major header.
        /// 
        /// Sometimes a cell needs a border on the bottom for styling, but instead of it having the border
        /// on the bottom, the cell below it has a border on top (and the same goes for all sides of the cell).
        /// For this reason, we cannot remove any borders if the adjacent cell on that side is not empty.
        /// </summary>
        /// <param name="worksheet">the worksheet being cleaned</param>
        /// <param name="firstDataRow">the first row that is part of the table</param>
        protected virtual void RemoveUnwantedBorders(ExcelWorksheet worksheet, int firstDataRow)
        {
            if (this.BorderRemoval == BorderRemovalType.NONE)
            {
                return;
            }


            ExcelRange cell;

            for(int row = 1; row < firstDataRow; row++)
            {
                for(int col = 1; col <= worksheet.Dimension.End.Column; col++)
                {
                    cell = worksheet.Cells[row, col];


                    if(this.BorderRemoval == BorderRemovalType.ONLY_EMPTY_CELLS)
                    {
                        RemoveCellBordersIfEmpty(worksheet, cell);
                    }
                    else
                    {
                        //Just remove all borders whether the cell is empty or not (except the last row)
                        bool lastRow = (row + 1 == firstDataRow);
                        ClearAllBorders(cell, lastRow);
                    }
                }
            }
        }




        /// <summary>
        /// Removes the border on each side of the specified cell, only if the cell is empty and so is the adjacent 
        /// cell on that side.
        /// </summary>
        /// <param name="worksheet">the worksheet being cleaned</param>
        /// <param name="cell">the cell whose borders should be deleted</param>
        protected virtual void RemoveCellBordersIfEmpty(ExcelWorksheet worksheet, ExcelRange cell)
        {
            ClearTopBorder(worksheet, cell);
            ClearBottomBorder(worksheet, cell);
            ClearLeftBorder(worksheet, cell);
            ClearRightBorder(worksheet, cell);
        }



        /// <summary>
        /// Removes the top border of the specified cell if appropriate.
        /// 
        /// Note: do not use this function if you want to remove ALL borders and not just those that are for empty cells.
        /// </summary>
        /// <param name="worksheet">the worksheet being cleaned</param>
        /// <param name="cell">the cell whose top border should be cleared</param>
        protected void ClearTopBorder(ExcelWorksheet worksheet, ExcelRange cell)
        {
            if (!IsEmptyCell(cell))
            {
                return;

            }
            if (cell.Style.Border.Top.Style == ExcelBorderStyle.None)
            {
                return;
            }
            if (cell.End.Row == 1) //if the adjcent cell is out of bounds
            {
                cell.Style.Border.Top.Style = ExcelBorderStyle.None;
                return;
            }



            ExcelRange adjacentCell = worksheet.Cells[cell.End.Row - 1, cell.End.Column]; //get cell above
            if (IsEmptyCell(adjacentCell))
            {
                cell.Style.Border.Top.Style = ExcelBorderStyle.None;
            }
        }



        /// <summary>
        /// Removes the bottom border of the specified cell if appropriate.
        /// 
        /// Note: do not use this function if you want to remove ALL borders and not just those that are for empty cells.
        /// </summary>
        /// <param name="worksheet">the worksheet being cleaned</param>
        /// <param name="cell">the cell whose bottom border should be cleared</param>
        protected void ClearBottomBorder(ExcelWorksheet worksheet, ExcelRange cell)
        {
            if (!IsEmptyCell(cell))
            {
                return;
            }
            if (cell.Style.Border.Bottom.Style == ExcelBorderStyle.None)
            {
                return;
            }
            if (cell.End.Row == worksheet.Dimension.End.Row) //if the adjcent cell is out of bounds
            {
                cell.Style.Border.Bottom.Style = ExcelBorderStyle.None;
                return;
            }



            ExcelRange adjacentCell = worksheet.Cells[cell.End.Row + 1, cell.End.Column]; //get cell below
            if (IsEmptyCell(adjacentCell))
            {
                cell.Style.Border.Bottom.Style = ExcelBorderStyle.None;
            }
        }



        /// <summary>
        /// Removes the left border of the specified cell if appropriate.
        /// 
        /// Note: do not use this function if you want to remove ALL borders and not just those that are for empty cells.
        /// </summary>
        /// <param name="worksheet">the worksheet being cleaned</param>
        /// <param name="cell">the cell whose left border should be cleared</param>
        protected void ClearLeftBorder(ExcelWorksheet worksheet, ExcelRange cell)
        {
            if (!IsEmptyCell(cell))
            {
                return;
            }
            if (cell.Style.Border.Left.Style == ExcelBorderStyle.None)
            {
                return;
            }
            if (cell.End.Column == 1) //if the adjcent cell is out of bounds
            {
                cell.Style.Border.Left.Style = ExcelBorderStyle.None;
                return;
            }



            ExcelRange adjacentCell = worksheet.Cells[cell.End.Row, cell.End.Column - 1]; //get cell o the left
            if (IsEmptyCell(adjacentCell))
            {
                cell.Style.Border.Left.Style = ExcelBorderStyle.None;
            }
        }




        /// <summary>
        /// Removes the right border of the specified cell if appropriate.
        /// 
        /// Note: do not use this function if you want to remove ALL borders and not just those that are for empty cells.
        /// </summary>
        /// <param name="worksheet">the worksheet being cleaned</param>
        /// <param name="cell">the cell whose right border should be cleared</param>
        protected void ClearRightBorder(ExcelWorksheet worksheet, ExcelRange cell)
        {
            if (!IsEmptyCell(cell))
            {
                return;
            }
            if (cell.Style.Border.Right.Style == ExcelBorderStyle.None)
            {
                return;
            }
            if (cell.End.Row == worksheet.Dimension.End.Column) //if the adjcent cell is out of bounds
            {
                cell.Style.Border.Right.Style = ExcelBorderStyle.None;
                return;
            }



            ExcelRange adjacentCell = worksheet.Cells[cell.End.Row, cell.End.Column + 1]; //get cell to the right
            if (IsEmptyCell(adjacentCell))
            {
                cell.Style.Border.Right.Style = ExcelBorderStyle.None;
            }
        }



        /// <summary>
        /// Removes all borders regaurdless of the cell contents except for the bottom border
        /// of the last row. That border is not removed becuase removing it might damage the display 
        /// of the data table on the row below.
        /// </summary>
        /// <param name="cell">the cell that needs its borders removed</param>
        /// <param name="lastRow">true if the specified cell is on the last row, and false otherwise</param>
        protected void ClearAllBorders(ExcelRange cell, bool lastRow)
        {
            //Just remove all borders whether the cell is empty or not (except the last row)
            cell.Style.Border.Top.Style = ExcelBorderStyle.None;
            cell.Style.Border.Left.Style = ExcelBorderStyle.None;
            cell.Style.Border.Right.Style = ExcelBorderStyle.None;


            if (!lastRow)
            {
                cell.Style.Border.Bottom.Style = ExcelBorderStyle.None;
            }
        }



        /// <summary>
        /// Finds the full ExcelRange object that contains the entire merge at the specified address. 
        /// In other words, the specified row and column point to a cell that is merged to be part of a 
        /// larger cell. This method returns the ExcelRange for the ENTIRE merge cell.
        /// </summary>
        /// <param name="worksheet">the worksheet we are currently cleaning</param>
        /// <param name="row">the row of a cell that is part of the larger merge</param>
        /// <param name="col">the column of a cell that is part of the larger merge</param>
        /// <returns>the Excel range object containing the entire merge, or null if the specifed cell is not a merge</returns>
        public static ExcelRange GetMergeCellByPosition(ExcelWorksheet worksheet, int row, int col)
        {
            int index = worksheet.GetMergeCellId(row, col);

            if (index < 1)
            {
                return null;
            }

            string cellAddress = worksheet.MergedCells[index - 1];
            return worksheet.Cells[cellAddress];
        }
    }
}

