using ExcelCleanerNet45;
using OfficeOpenXml;
using System;
using System.CodeDom;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelCleanerNet45.GeneralCleaning
{

    /// <summary>
    /// An extentsion of the primary merge cleaner that adds realingment of data cells to the additional cleanup.
    /// 
    /// The realignment consists of deleteing all empty cells inside the data section of the report and setting the
    /// of each column to a minimum value.
    /// </summary>
    internal class ReAlignDataCells : PrimaryMergeCleaner
    {
        private static readonly double MIN_COLUM_WIDTH = 10.3;
        private string bottomHeader;



        /// <summary>
        /// 
        /// </summary>
        /// <param name="bottomHeader">the text that marks the bottom row of the data section, beyond which no 
        /// realignment should be done. If null, or not provided, will default to the last row in the worksheet</param>
        public ReAlignDataCells(string bottomHeader = null)
        {
            this.bottomHeader = bottomHeader;
        }



        /// <inheritdoc/>
        protected override void AdditionalCleanup(ExcelWorksheet worksheet)
        {
            base.AdditionalCleanup(worksheet);
            
            RealignCells(worksheet);
        }



        /// <summary>
        /// Re-aligns cells by deleteing all empty cells in the data section of the table
        /// </summary>
        /// <param name="worksheet">the worksheet currently being cleaned</param>
        protected void RealignCells(ExcelWorksheet worksheet)
        {
            //Get boundries of the section that needs to be re-aligned
            int startRow = base.firstRowOfTable;
            int endRow = FindBottomRow(worksheet);
            int startCol = FindStartColumn(worksheet);
            int endCol = worksheet.Dimension.End.Column;


            SetColumnWidths(worksheet, startCol, endCol);

            DeleteEmptyCells(worksheet, startRow, endRow, startCol, endCol);
            
        }



        /// <summary>
        /// Finds the row, boyond which no re-alignment should be done. This row is found based on the text passed to the
        /// constructor of this class.
        /// </summary>
        /// <param name="worksheet">the worksheet being cleaned</param>
        /// <returns>the row number of the last row that is part of the data section</returns>
        protected int FindBottomRow(ExcelWorksheet worksheet)
        {
            if(bottomHeader == null)
            {
                return worksheet.Dimension.End.Row;
            }


            ExcelIterator iter = new ExcelIterator(worksheet);

            ExcelRange cell = iter.GetFirstMatchingCell(c => FormulaManager.TextMatches(c.Text, bottomHeader));

            return cell.End.Row;
        }




        /// <summary>
        /// Finds the first (leftmost) column in the worksheet that contains numeric data.
        /// </summary>
        /// <param name="worksheet">the worksheet currently being cleaned</param>
        /// <returns>the column number of the first (leftmost) column that contains numeric data</returns>
        protected int FindStartColumn(ExcelWorksheet worksheet)
        {
            ExcelRange cell;

            for(int col = 1; col <= worksheet.Dimension.End.Column; col++)
            {
                for(int row = base.firstRowOfTable; row <= worksheet.Dimension.End.Row; row++)
                {
                    cell = worksheet.Cells[row, col];

                    if (HasNumericData(cell))
                    {
                        return col;
                    }
                }
            }


            return 1; //default
        }




        /// <summary>
        /// Checks if the specified cell contains numeric data
        /// </summary>
        /// <param name="cell">the cell being checked</param>
        /// <returns>true if the cell contains numeric data and false otherwise</returns>
        protected bool HasNumericData(ExcelRange cell)
        {
            double ignored;

            return cell.Text.StartsWith("$") || cell.Text.StartsWith("%") || Double.TryParse(cell.Text, out ignored);
        }




        /// <summary>
        /// Sets the width of all columns between the specified column numbers to be no less then MIN_COLUM_WIDTH.
        /// </summary>
        /// <param name="worksheet">the worksheet currently being cleaned</param>
        /// <param name="startCol">the first column that needs resizing</param>
        /// <param name="endCol">the last column that needs resizing</param>
        protected void SetColumnWidths(ExcelWorksheet worksheet, int startCol, int endCol)
        {
            ExcelColumn currentCol;

            for(int i = startCol; i <= endCol; i++)
            {
                currentCol = worksheet.Column(i);

                if(currentCol.Width < MIN_COLUM_WIDTH)
                {
                    currentCol.Width = MIN_COLUM_WIDTH;
                }
            }
        }



        /// <summary>
        /// Deletes all empty cells inside the specified area (inclusive)
        /// </summary>
        /// <param name="worksheet">the worksheet being cleaned</param>
        /// <param name="startRow">the first row in the area</param>
        /// <param name="endRow">the last row in the area</param>
        /// <param name="startCol">the first column in the area</param>
        /// <param name="endCol">the last column in the area</param>
        protected void DeleteEmptyCells(ExcelWorksheet worksheet, int startRow, int endRow, int startCol, int endCol)
        {
            ExcelRange cell;

            for (int col = endCol; col >= startCol; col--)
            {
                for (int row = startRow; row <= endRow; row++)
                {
                    cell = worksheet.Cells[row, col];

                    if (base.IsEmptyCell(cell))
                    {
                        cell.Delete(eShiftTypeDelete.Left);
                    }
                }
            }
        }
    }
}
