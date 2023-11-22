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


            //TODO: add formulas to the inner and outer segments
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
    }
}
