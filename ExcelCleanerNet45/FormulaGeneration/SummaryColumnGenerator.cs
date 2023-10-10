using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ExcelCleanerNet45.FormulaGeneration
{

    public delegate bool IsEndOfColumn(ExcelRange cell);


    /// <summary>
    /// Implementation of IFormulaGenerator that makes every data cell in a column into the sum of the corrisponding
    /// data cells in other columns. The column headers should be passed to this class in this format: 
    /// 
    /// col1,col2,col3=>summaryCol. 
    /// 
    /// Where coln and summaryCol should be regexes that match to the headers of the column needed.
    /// 
    /// Note: only 1 column is required on the left side, more than one is optional.
    /// </summary>
    internal class SummaryColumnGenerator : IFormulaGenerator
    {

        private IsDataCell dataCellDef = new IsDataCell(FormulaManager.IsDollarValue);
        private IsEndOfColumn endOfColumnDef = new IsEndOfColumn(cell => cell.Style.Font.Bold);



        public void InsertFormulas(ExcelWorksheet worksheet, string[] headers)
        {
            foreach(string header in headers)
            {
                if (!MatchesHeaderFormat(header))
                {
                    continue;
                }


                //Find the desired header texts
                string summaryHeader = header.Substring(header.IndexOf("=>") + 2);
                string[] dataColHeaders = header.Substring(0, header.IndexOf("=>")).Split(',');
                bool[] isColNegative = CheckColumnNegativity(dataColHeaders);



                //Find header of summary column
                Tuple<int, int> summaryCoords = FindColumnStartCoordinates(worksheet, summaryHeader);
                int row = summaryCoords.Item1;
                int summaryColumn = summaryCoords.Item2;



                Dictionary<int, bool> columns = FindAllDataColumns(worksheet, row, dataColHeaders, isColNegative);

                AddFormulaToEachCell(worksheet, columns, row, summaryColumn);
            }
        }




        /// <summary>
        /// Checks if the specified header matches the format that is required by this implementation of IFormulaGenerator.
        /// If it does not match the required format, it is most likly intended for another forula generator.
        /// </summary>
        /// <param name="header">the header being checked</param>
        /// <returns>true if the header matches the required format, and false otherwise</returns>
        private bool MatchesHeaderFormat(string header)
        {
            return header.IndexOf("=>") >= 0;
        }




        /// <summary>
        /// Strips the negative sign off the beginning of each header and returns an array with a true for 
        /// each header that had a negative sign removed, and a false for the rest.
        /// </summary>
        /// <param name="headers">the array of headers that should be moved to a dictionary</param>
        /// <returns>an array of booleans representing isNegative for each header</returns>
        private bool[] CheckColumnNegativity(string[] headers)
        {
            bool[] isNegative = new bool[headers.Length];

            for(int i = 0; i < headers.Length; i++)
            {
                if (headers[i].StartsWith("-"))
                {
                    headers[i] = headers[i].Substring(1);
                    isNegative[i] = true;
                }
                else
                {
                    isNegative[i] = false;
                }
            }

            return isNegative;
        }




        /// <summary>
        /// Finds the coordinates of the first cell in the worksheet with the specified text in it
        /// </summary>
        /// <param name="worksheet">the worksheet currently being given formulas</param>
        /// <param name="headerText">the text to search for</param>
        /// <returns>the coordinates of the first cell with text matching the specified text</returns>
        private Tuple<int, int> FindColumnStartCoordinates(ExcelWorksheet worksheet, string headerText)
        {
            ExcelIterator iter = new ExcelIterator(worksheet);

            return iter.FindAllMatchingCoordinates(cell => FormulaManager.TextMatches(cell.Text, headerText)).First();
        }




        /// <summary>
        /// Finds all columns that have one of the specified headers and adds its column number to a dictionary along with a
        /// bool isNegative.
        /// </summary>
        /// <param name="worksheet">the worksheet currently being given formulas</param>
        /// <param name="startingRow">the row where the column headers are found</param>
        /// <param name="headers">the headers to look for to detect data columns</param>
        /// <param name="isColNegative">an array of isNegative booleans corisponding to each header</param>
        /// <returns>a dictionary of all column numbers and a true if they are negative</returns>
        private Dictionary<int, bool> FindAllDataColumns(ExcelWorksheet worksheet, int startingRow, string[] headers, bool[] isColNegative)
        {
            Dictionary<int, bool> columns = new Dictionary<int, bool>();



            ExcelIterator iter = new ExcelIterator(worksheet, startingRow, 1);
            var coordsInRow = iter.GetCellCoordinates(ExcelIterator.SHIFT_RIGHT);
            foreach(Tuple<int, int> coords in coordsInRow)
            {

                ExcelRange cell = worksheet.Cells[coords.Item1, coords.Item2];

                if ((FormulaManager.IsEmptyCell(cell)))
                {
                    continue;
                }


                for(int i = 0; i < headers.Length; i++)
                {
                    if (FormulaManager.TextMatches(cell.Text, headers[i]))
                    {
                        columns.Add(coords.Item2, isColNegative[i]);
                        break;
                    }
                }
            }


            return columns;
        }




        /// <summary>
        /// Iterates through every cell in the summary column and adds the appropriate formula to it
        /// </summary>
        /// <param name="worksheet">the worksheet in need of formulas</param>
        /// <param name="otherCols">a dictionary with each column number to be included in the formula, and a true if they should be subtracted</param>
        /// <param name="headerRow">the row number of the headers of each column</param>
        /// <param name="col">the column number of the summary column</param>
        private void AddFormulaToEachCell(ExcelWorksheet worksheet, Dictionary<int, bool> otherCols, int headerRow, int col)
        {
            ExcelIterator iter = new ExcelIterator(worksheet, headerRow + 1, col);
            var eachCell = iter.GetCells(ExcelIterator.SHIFT_DOWN, (cell => this.endOfColumnDef(cell)));
            foreach(ExcelRange cell in eachCell)
            {
                if (dataCellDef(cell))
                {
                    cell.Formula = BuildFormula(worksheet, otherCols, cell.Start.Row);
                    cell.Style.Locked = true;
                }
            }
        }




        /// <summary>
        /// Builds the formula that should be added to the cell
        /// </summary>
        /// <param name="worksheet">the worksheet being given formulas</param>
        /// <param name="otherCols">each other column number that should be included in the formula, and if it should be subtracted</param>
        /// <param name="row">the row we are currently on</param>
        /// <returns>a formula to use in the report</returns>
        private string BuildFormula(ExcelWorksheet worksheet, Dictionary<int, bool> otherCols, int row)
        {
            StringBuilder result = new StringBuilder("SUM(");


            ExcelRange cell;
            foreach(int colNumber in otherCols.Keys)
            {
                cell = worksheet.Cells[row, colNumber];

                // if column should be subtracted
                if (otherCols[colNumber])
                {
                    result.Append("-");
                }

                result.Append(cell.Address);
                result.Append(",");
            }


            result.Remove(result.Length - 1, 1);
            result.Append(")");

            return result.ToString();
        }




        public void SetDataCellDefenition(IsDataCell isDataCell)
        {
            this.dataCellDef = isDataCell;
        }



        public void SetEndOfColumnDefenition(IsEndOfColumn isEndOfCol)
        {
            this.endOfColumnDef = isEndOfCol;
        }
    }
}
