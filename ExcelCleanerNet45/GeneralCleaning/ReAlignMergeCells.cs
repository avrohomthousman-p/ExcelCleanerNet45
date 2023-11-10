using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelCleanerNet45.GeneralCleaning
{

    /// <summary>
    /// An extentsion of the primary merge cleaner ensures that all data cells (and the header) in a given column all
    /// have a merge that spans the same range.
    /// 
    /// Some reports (like ProfitAndLossBudget) have most of their data cells of a column merged over the same area, but
    /// then some cells (usually the column header) merged over a different area. This causes the column to not be aligned
    /// correctly after running the primary merge cleaner. The purpose of this class is to correct that issue before 
    /// the primary merge cleaner begins. After this class does its cleanup, the primary merge cleaner should be able
    /// to do its work as though this were a regular report.
    /// </summary>
    class ReAlignMergeCells : PrimaryMergeCleaner
    {
        public override void Unmerge(ExcelWorksheet worksheet)
        {
            FindTableBounds(worksheet);

            ReAlignWorksheet(worksheet);

            UnMergeMergedSections(worksheet);

            ResizeCells(worksheet);

            DeleteColumns(worksheet);

            AdditionalCleanup(worksheet);
        }




        /// <summary>
        /// Executes the reAlignment of each data column in the worksheet
        /// </summary>
        /// <param name="worksheet"></param>
        private void ReAlignWorksheet(ExcelWorksheet worksheet)
        {
            foreach(Tuple<int, int> range in mergeRangesOfDataCells)
            {
                Tuple<int, int> desiredColumnSpan = GetProperMergeRange(worksheet, range);

                SetAllMergeSpans(worksheet, range, desiredColumnSpan);
            }
        }



        /// <summary>
        /// Scans all non-empty cells inside the column range and determans the mode average merge column span.
        /// (which columns the majority of merge cells span)
        /// </summary>
        /// <param name="worksheet">the worksheet being cleaned</param>
        /// <param name="header">the columns that the column header spans</param>
        /// <returns>A merge range that all cells in the column should span</returns>
        private Tuple<int, int> GetProperMergeRange(ExcelWorksheet worksheet, Tuple<int, int> header)
        {
            //tracks each span of a merge cell to how many times a cell like that was found
            Dictionary<Tuple<int, int>, int> mergeSpans = new Dictionary<Tuple<int, int>, int>();

            for(int i = base.firstRowOfTable; i <= worksheet.Dimension.End.Row; i++)
            {

                //for each column below our headers
                for(int j = header.Item1; j <= header.Item2; j++)
                {
                    ExcelRange cell = GetMergeCellByPosition(worksheet, i, j);

                    if(cell == null || IsEmptyCell(cell))
                    {
                        continue;
                    }
                    else
                    {
                        Tuple<int, int> span = new Tuple<int, int>(cell.Start.Column, cell.End.Column);

                        //update dictionary
                        if (!mergeSpans.ContainsKey(span))
                        {
                            mergeSpans.Add(span, 0);
                        }
                        mergeSpans[span]++;


                        j = cell.End.Column;
                    }
                }
            }



            return GetKeyWithLargestValue(mergeSpans);
        }




        /// <summary>
        /// Searches the specified dictionary and finds and returns the key that has the largest value 
        /// associated with it.
        /// </summary>
        /// <param name="dictionary">the dictionary to be searched</param>
        /// <returns>the key with the largest associated value</returns>
        private Tuple<int, int> GetKeyWithLargestValue(Dictionary<Tuple<int, int>, int> dictionary)
        {
            Tuple<int, int> key = null;
            int largestValue = Int32.MinValue;
            
            foreach(KeyValuePair<Tuple<int, int>, int> kvp in dictionary)
            {
                if(kvp.Value >= largestValue)
                {
                    largestValue = kvp.Value;
                    key = kvp.Key;
                }
            }


            return key;
        }



        /// <summary>
        /// Sets all the merge cells in the original column span, to only cover the desired column span
        /// </summary>
        /// <param name="worksheet">the worksheet currently being cleaned</param>
        /// <param name="originalSpan">the span of the header of the data column</param>
        /// <param name="desiredSpan">the columns every data cell in the column SHOULD span</param>
        private void SetAllMergeSpans(ExcelWorksheet worksheet, Tuple<int, int> originalSpan, Tuple<int, int> desiredSpan)
        {
            //TODO: iterate through the column and ensure all merge cellsmatch the desired span
        }

    }
}
