using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelCleanerNet45.GeneralCleaning
{

    /// <summary>
    /// An override of the PrimaryMergeCleaner that ensures that each data column in the worksheet
    /// has a minimum width.
    /// </summary>
    internal class SetDefaultColumnWidth : PrimaryMergeCleaner
    {
        protected readonly double DEFAULT_COLUMN_WIDTH = 11;



        /// <inheritdoc/>
        protected override void ResizeCells(ExcelWorksheet worksheet)
        {
            base.ResizeCells(worksheet);


            ExcelIterator iter = new ExcelIterator(worksheet, base.firstRowOfTable, 1);
            var dataCols = iter.GetCells(ExcelIterator.SHIFT_RIGHT)
                                .Where(cell => !IsEmptyCell(cell))
                                .Select(cell => cell.Start.Column);


            foreach (int columnNum in dataCols)
            {
                var column = worksheet.Column(columnNum);

                if (column.Hidden)
                {
                    column.Hidden = false;
                }

                
                if (column.Width < DEFAULT_COLUMN_WIDTH)
                {
                    column.Width = DEFAULT_COLUMN_WIDTH;
                }
                
            }

        }
    }
}
