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


            //if there is no second column, this report can be processed by the regular RowSegmentFormulaGenerator
            if(secondCol == -1)
            {
                Console.WriteLine("Only one column has headers. Switching to the RowSegmentFormulaGenerator");
                var generator = new RowSegmentFormulaGenerator();
                generator.trimFormulaRange = this.trimFormulaRange;
                generator.SetDataCellDefenition(this.isDataCell);
                generator.InsertFormulas(worksheet, headers);
                return;
            }


            DoOuterSegments(worksheet, firstCol, secondCol);
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
        /// Adds the formulas to the outer segments and does the function calls needed to do the
        /// inner segments.
        /// </summary>
        /// <param name="worksheet">the worksheet in need of formulas</param>
        /// <param name="mainColumn">the column containing the headers of the outer segments</param>
        /// <param name="innerColumn">the column containing the headers of the inner segments</param>
        protected virtual void DoOuterSegments(ExcelWorksheet worksheet, int mainColumn, int innerColumn)
        {
            foreach(var range in GetAllFormulaSegments(worksheet, mainColumn))
            {
                //If there is an inner range within these rows and in the innerColumn
                //we should do that range first, and give this range a formula that
                //adds up only the formulas in the inner range
                Tuple<int, int>[] innerSegments = GetAllFormulaSegments(worksheet, innerColumn, 
                                                                        range.Item1, range.Item2)
                                                                        .ToArray();


                bool hasInnerSegments = DoInnerSegment(worksheet, innerSegments, innerColumn);
                if(hasInnerSegments)
                {
                    //build an array formula for the outer segment
                    FillInSumOfSumsFormulas(worksheet, range.Item1, range.Item2, mainColumn);
                }
                else //Otherwise, just do regular formulas
                {
                    FillInRegularFormulas(worksheet, range.Item1, range.Item2, mainColumn);
                }
            }
        }



        /// <summary>
        /// Returns all the Formula Segments in the table if they are in the specified column as a Tuple of 
        /// [startRow, endRow]
        /// </summary>
        /// <param name="worksheet">the worksheet in need of formulas</param>
        /// <param name="col">the column to look for headers in</param>
        /// <returns>the formula ranges as tuples of [start row, end row]</returns>
        protected IEnumerable<Tuple<int, int>> GetAllFormulaSegments(ExcelWorksheet worksheet, int col)
        {
            return GetAllFormulaSegments(worksheet, col, 1, worksheet.Dimension.End.Row);
        }




        /// <summary>
        /// Returns all the Formula Segments that are in the specified column and between the specified top and
        /// bottom row (incluseive) as a Tuple of [startRow, endRow]
        /// </summary>
        /// <param name="worksheet">the worksheet in need of formulas</param>
        /// <param name="col">the column to look for headers in</param>
        /// <param name="topRow">to top most row that should be checked</param>
        /// <param name="bottomRow">the bottom most row that should be checked</param>
        /// <returns>the formula ranges as tuples of [start row, end row]</returns>
        protected virtual IEnumerable<Tuple<int, int>> GetAllFormulaSegments(ExcelWorksheet worksheet, int col, int topRow, int bottomRow)
        {
            //if we got a value that is out of bounds, just set it to the first/last row of the worksheet
            topRow = Math.Max(topRow, 1);
            bottomRow = Math.Min(bottomRow, worksheet.Dimension.End.Row);



            int indexOfStartHeader;
            ExcelRange cell;
            ExcelIterator iter = new ExcelIterator(worksheet, topRow, col);

            for (int i = topRow; i <= bottomRow; i++)
            {
                cell = worksheet.Cells[i, col];

                indexOfStartHeader = GetMatchingElement(startHeaders, cell.Text);

                if (indexOfStartHeader == -1) //it doesnt match any headers
                {
                    continue;
                }



                int endRow = FindMatchingEndHeader(worksheet, cell.Start.Row + 1,
                                                    col, endHeaders[indexOfStartHeader]);

                if (endRow == -1 || endRow > bottomRow)
                {
                    continue;
                }


                //FillInInnerFormulas(worksheet, cell.Start.Row, endRow, col);
                yield return new Tuple<int, int>(cell.Start.Row, endRow);


                //advance iterator to end of this segment (if doing so wont cause OOB error)
                if (endRow < worksheet.Dimension.End.Row)
                {
                    iter.SetCurrentLocation(endRow, col);
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
        protected virtual int FindMatchingEndHeader(ExcelWorksheet worksheet, int startRow, int col, string endHeaderRegex)
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
        /// Adds all the neccisary formulas to all formula ranges passed in. Returns true if there was an inner
        /// segment to give formulas to, and false if there was none.
        /// </summary>
        /// <param name="worksheet">the worksheet in need of formulas</param>
        /// <param name="formulaRanges">an array of tuples containing the start and end column of each formula range</param>
        /// <param name="column">the column we are checking in</param>
        /// <returns>true if at least one formula was added and false otherwise</returns>
        protected virtual bool DoInnerSegment(ExcelWorksheet worksheet, Tuple<int, int>[] formulaRanges, int column)
        {
            if (formulaRanges == null || formulaRanges.Length == 0)
            {
                return false;
            }


            //add the formulas to each range
            foreach (Tuple<int, int> range in formulaRanges)
            {
                FillInRegularFormulas(worksheet, range.Item1, range.Item2, column);
            }

            return true;
        }




        /// <summary>
        /// Inserts regular (non-array) formulas in each cell in the formula range that requires it.
        /// </summary>
        /// <param name="worksheet">the worksheet currently being given formulas</param>
        /// <param name="startRow">the first row of the formula range (containing the header)</param>
        /// <param name="endRow">the last row of the formula range (containing the total)</param>
        /// <param name="col">the column of the header and total for the formula range</param>
        protected virtual void FillInRegularFormulas(ExcelWorksheet worksheet, int startRow, int endRow, int col)
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




        /// <summary>
        /// Inserts formulas in the specified formula range that add up all the other formulas in the range (from inner segments)
        /// </summary>
        /// <param name="worksheet">the worksheet currently being given formulas</param>
        /// <param name="startRow">the first row of the formula range (containing the header)</param>
        /// <param name="endRow">the last row of the formula range (containing the total)</param>
        /// <param name="col">the column of the header and total for the formula range</param>
        protected virtual void FillInSumOfSumsFormulas(ExcelWorksheet worksheet, int startRow, int endRow, int col)
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



                    ExcelRange range = worksheet.Cells[startRow, col, endRow - 1, col];

                    //Formula to add up all cells that contain a formula
                    //The _xlfn fixes a bug in excel
                    string formula = $"SUM(IF(_xlfn.ISFORMULA({range.Address}), {range.Address}, 0))";

                    cell.CreateArrayFormula(formula);
                    cell.Style.Locked = true;
                    cell.Style.Hidden = false;
                    cell.Calculate();

                }
                else if (!FormulaManager.IsEmptyCell(cell))
                {
                    return;
                }
            }
        }
    }
}
