using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace ExcelCleanerNet45.FormulaGeneration.ReportSpecificGenerators
{
    /// <summary>
    /// Implementation of IFormulaGenerator that is designed specifically for the TrialBalanceVariance report.
    /// The TrialBalanceVarience report is strutured like all reports that need the RowSegmentFormulaGenerator
    /// except that it has "formula segments" (parts that need their own formulas like Assets -> Total Assets)
    /// inside other formula segments. This class gives the inner segments their formulas as the 
    /// RowSegmentFormulaGenerator would, and gives the outer segments formulas that only sum up the totals
    /// inside their ranges.
    /// 
    /// Arguments for this class should be the same as the RowSegmentFormulaGenerator for both inner and outer segments.
    /// </summary>
    internal class TrialBalanceVarianceGenerator : IFormulaGenerator
    {
        protected List<string> startHeaders;
        protected List<string> endHeaders;


        protected IsDataCell isDataCell = new IsDataCell(FormulaManager.IsDollarValue);


        private bool trimRange = false;
        /// <summary>
        /// controls whether or not the system skips the whitespace immideatly after the top header of each formula range
        /// e.g. if we are summing up all cells from "Income" to "Total Income", this boolean decides if the formula 
        /// should include the empty cells just below the "Income" row
        /// </summary>
        public bool trimFormulaRange
        {
            get { return trimRange; }
            set { trimRange = value; }
        }



        public void SetDataCellDefenition(IsDataCell isDataCell)
        {
            this.isDataCell = isDataCell;
        }




        public void InsertFormulas(ExcelWorksheet worksheet, string[] headers)
        {
            BuildHeaderList(headers);


            //These are the columns where the headers will be found
            int firstCol = FindNextColumnWithHeaders(worksheet, 1);
            int secondCol = FindNextColumnWithHeaders(worksheet, firstCol + 1);


            DoInnerSegments(worksheet, secondCol);
            //TODO: add formulas to the segments
        }




        /// <summary>
        /// Converts the header arguments given to this class into two seperate string arrays,
        /// one with the start headers and one with the end headers.
        /// </summary>
        /// <param name="headers"></param>
        protected virtual void BuildHeaderList(string[] headers)
        {
            startHeaders = new List<string>(headers.Length);
            endHeaders = new List<string>(headers.Length);
            string currentHeader;

            for(int i = 0; i < headers.Length; i++)
            {
                currentHeader = headers[i];

                //Ensure that the header was intended for this class and not the DistantRowsFormulaGenerator class
                if (FormulaManager.IsNonContiguousFormulaRange(currentHeader))
                {
                    continue;
                }


                int seperator = currentHeader.IndexOf('=');

                startHeaders.Add(currentHeader.Substring(0, seperator));
                endHeaders.Add(currentHeader.Substring(seperator + 1));
            }
        }




        /// <summary>
        /// Finds the first column that contains at least one start header and its matching end header and does
        /// not appear before the specified start column.
        /// </summary>
        /// <param name="worksheet">the worksheet in need of formulas</param>
        /// <param name="startColumn">the column our search should start at. Columns before this will not be checked</param>
        /// <returns>
        /// the column number of the first column found that contains a matching start and end header, 
        /// or -1 if no such column can be found
        /// </returns>
        protected virtual int FindNextColumnWithHeaders(ExcelWorksheet worksheet, int startColumn)
        {
            for(int col = startColumn; col <= worksheet.Dimension.End.Column; col++)
            {
                if(CheckColumnForHeaders(worksheet, col))
                {
                    return col;
                }
            }


            return -1;
        }




        /// <summary>
        /// Checks if the specified column contains any start header and the corrisponding end header
        /// </summary>
        /// <param name="worksheet">the worksheet in need of formulas</param>
        /// <param name="column">the column being checked</param>
        /// <returns>true if the column has a matching start and end header and false otherwise</returns>
        protected virtual bool CheckColumnForHeaders(ExcelWorksheet worksheet, int column)
        {
            ExcelIterator iter = new ExcelIterator(worksheet, 1, column);


            //first find start header
            foreach(ExcelRange cell in iter.GetCells(ExcelIterator.SHIFT_DOWN))
            {
                if(FormulaManager.IsDollarValue(cell) || FormulaManager.IsEmptyCell(cell))
                {
                    continue;
                }


                //check if the current cell matches any start headers
                int matchedStartHeader = GetMatchingElement(startHeaders, cell.Text);

                if(matchedStartHeader != -1) //if we found a match
                {
                    if(HasMatchingEndHeader(iter, endHeaders[matchedStartHeader]))
                    {
                        return true;
                    }
                }
            }


            return false;
        }




        /// <summary>
        /// Gets the index of the specified array that contains a regex matching the specified text.
        /// </summary>
        /// <param name="regexes">a list of regexes to be compared to the text</param>
        /// <param name="text">the text to compare to each regex</param>
        /// <returns>the index containing the regex that matches the specified text, or -1 if no matches are found</returns>
        protected int GetMatchingElement(List<string> regexes, string text)
        {
            for (int i = 0; i < regexes.Count; i++)
            {
                if(FormulaManager.TextMatches(text, regexes[i]))
                {
                    return i;
                }
            }

            return -1;
        }



        /// <summary>
        /// Checks if any of the cells from the iterators position or below match the specified regex.
        /// </summary>
        /// <param name="iter">the iterator that references where we are holding in the column</param>
        /// <param name="endHeaderRegex">the regex our cell must match</param>
        /// <returns>true if a cell is found matching the regex, and false otherwise</returns>
        protected bool HasMatchingEndHeader(ExcelIterator iter, string endHeaderRegex)
        {
            ExcelIterator copy = new ExcelIterator(iter); //ensure we don't mess up the original iterator

            return copy.GetCells(ExcelIterator.SHIFT_DOWN)
                       .Any(cell => FormulaManager.TextMatches(cell.Text, endHeaderRegex));
        }



        /// <summary>
        /// Adds formulas to all the segments in the specified row. This function adds regular formulas,
        /// not the array formulas that are needed for the outer segments.
        /// </summary>
        /// <param name="worksheet">the worksheet in need of formulas</param>
        /// <param name="columnWithHeaders">the column the headers will be found in</param>
        protected virtual void DoInnerSegments(ExcelWorksheet worksheet, int columnWithHeaders)
        {
            int matchingStartHeader;
            ExcelIterator iter = new ExcelIterator(worksheet, 1, columnWithHeaders);

            foreach(ExcelRange cell in iter.GetCells(ExcelIterator.SHIFT_DOWN))
            {
                matchingStartHeader = GetMatchingElement(startHeaders, cell.Text);

                if (matchingStartHeader == -1) //it doesnt match any headers
                {
                    continue;
                }
                    


                int endRow = FindRowOfEndHeader(worksheet, cell.Start.Row + 1,
                                                    columnWithHeaders, endHeaders[matchingStartHeader]);

                if (endRow == -1)
                {
                    continue;
                }


                FillInInnerFormulas(worksheet, cell.Start.Row, endRow, columnWithHeaders);


                //advance iterator to end of this segment
                if(endRow < worksheet.Dimension.End.Row)
                {
                    iter.SetCurrentLocation(endRow, columnWithHeaders);
                }
            }
        }



        /// <summary>
        /// Finds the row that has the end header matching the specifed regex
        /// </summary>
        /// <param name="worksheet">the worksheet in need of formulas</param>
        /// <param name="startRow">the first row to be checked for the end header</param>
        /// <param name="col">the column to check in</param>
        /// <param name="endHeaderRegex">the regex that the desired end header will match</param>
        /// <returns>the row number of the end header, or -1 if no end header is found</returns>
        protected virtual int FindRowOfEndHeader(ExcelWorksheet worksheet, int startRow, int col, string endHeaderRegex)
        {
            ExcelIterator iter = new ExcelIterator(worksheet, startRow, col);

            ExcelRange endCell = iter.GetCells(ExcelIterator.SHIFT_DOWN)
                                    .FirstOrDefault(cell => FormulaManager.TextMatches(cell.Text, endHeaderRegex));


            if(endCell == null)
            {
                return -1;
            }
            else
            {
                return endCell.End.Row;
            }
        }



        /// <summary>
        /// Inserts the formulas in each cell in the formula range that requires it.
        /// </summary>
        /// <param name="worksheet">the worksheet currently being given formulas</param>
        /// <param name="startRow">the first row of the formula range (containing the header)</param>
        /// <param name="endRow">the last row of the formula range (containing the total)</param>
        /// <param name="col">the column of the header and total for the formula range</param>
        protected virtual void FillInInnerFormulas(ExcelWorksheet worksheet, int startRow, int endRow, int col)
        {

            ExcelRange cell;



            //Often there are multiple columns that require a formula, so we need to iterate
            //and apply the formulas in many columns
            for (col++; col <= worksheet.Dimension.End.Column; col++)
            {
                cell = worksheet.Cells[endRow, col];

                if (this.isDataCell(cell))
                {
                    if (trimRange)
                    {
                        startRow += CountEmptyCellsOnTop(worksheet, startRow, endRow, col); //Skip the whitespace on top
                    }
                    string formula = FormulaManager.GenerateFormula(worksheet, startRow, endRow - 1, col);

                    FormulaManager.PutFormulaInCell(cell, formula);
                }
                else if (!FormulaManager.IsEmptyCell(cell))
                {
                    return;
                }
            }

        }



        /// <summary>
        /// Counts the number of empty cells between the start header(inclusive) and the actual data cells in the 
        /// formula range.
        /// </summary>
        /// <param name="worksheet">the worksheet in need of formulas</param>
        /// <param name="startRow">the row where the start header was found</param>
        /// <param name="endRow">te row where the end header was found</param>
        /// <param name="col">the column of the formula range</param>
        /// <returns>the number of empty cells at the start of the formula range</returns>
        protected static int CountEmptyCellsOnTop(ExcelWorksheet worksheet, int startRow, int endRow, int col)
        {
            int emptyCells = 0;
            ExcelRange cell;

            for (; startRow <= endRow; startRow++)
            {
                cell = worksheet.Cells[startRow, col];
                if (FormulaManager.IsEmptyCell(cell))
                {
                    emptyCells++;
                }
                else
                {
                    break;
                }
            }


            return emptyCells;
        }



        protected virtual void DoOuterSegments(ExcelWorksheet worksheet, int column)
        {
            //TODO
        }
    }
}
