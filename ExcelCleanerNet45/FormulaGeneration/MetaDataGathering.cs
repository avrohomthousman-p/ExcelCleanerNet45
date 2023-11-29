using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelCleanerNet45.FormulaGeneration
{
    /// <summary>
    /// Contains methods used to gather metadata about a report at runtime.
    /// </summary>
    internal static class MetaDataGathering
    {

        /// <summary>
        /// Gets the appropriate headers that are found in the PayablesAccountReport and are needed by the
        /// formula generator.
        /// </summary>
        /// <param name="worksheet">the worksheet that will soon be given formulas</param>
        /// <returns>an array of string arguments to be used by the formula generator of the PayablesAccountReport</returns>
        internal static string[] GetHeadersForPayablesAccountReport(ExcelWorksheet worksheet)
        {
            List<string> headers = new List<string>();


            int column = 1;
            var outerColumn = GetHeadersOfNextColumn(worksheet, ref column);
            column++;        //skips the column we just found
            var innerColumn = GetHeadersOfNextColumn(worksheet, ref column);



            //the multiformula generator requires each header to start with a number
            //indicating which formula generator its for. In this report, the outer
            //column gets the larger number
            outerColumn = outerColumn.Select(text => "2" + text);
            innerColumn = innerColumn.Select(text => "1" + text);


            headers.AddRange(innerColumn);
            headers.AddRange(outerColumn);



            headers.AddRange(GetFinalSummaryHeaders(outerColumn));


            return headers.ToArray();
        }




        /// <summary>
        /// Finds the first column that contains headers in it and returns those headers
        /// </summary>
        /// <param name="worksheet">the worksheet that needs formulas</param>
        /// <param name="startColumn">the column we should start the search at</param>
        /// <returns>all the header arguments required for the first column found, or null if no column is found</returns>
        private static IEnumerable<string> GetHeadersOfNextColumn(ExcelWorksheet worksheet, ref int startColumn)
        {
            IEnumerable<string> headers;

            for(; startColumn <= worksheet.Dimension.End.Column; startColumn++)
            {
                headers = GetAllHeadersInColumn(worksheet, startColumn);

                if(headers.Count() > 0)
                {
                    return headers;
                }
            }


            return null;
        }




        /// <summary>
        /// Gets all header arguments that cover all header start-end pairs in the specified column
        /// </summary>
        /// <param name="worksheet">the worksheet that will soon get formulas</param>
        /// <param name="column">the column we are scanning for start and end headers</param>
        /// <returns>each string argument that the formula manager will need for each header in this column</returns>
        private static IEnumerable<string> GetAllHeadersInColumn(ExcelWorksheet worksheet, int column)
        {
            ExcelIterator iter = new ExcelIterator(worksheet, 1, column);

            foreach (ExcelRange startCell in iter.GetCells(ExcelIterator.SHIFT_DOWN))
            {
                if (FormulaManager.IsEmptyCell(startCell) || !startCell.Style.Font.Bold)
                {
                    continue;
                }


                ExcelRange endCell = FindEndHeader(worksheet, startCell);
                if (endCell != null && endCell.Style.Font.Bold)
                {
                    yield return startCell.Text.Trim() + "=" + endCell.Text.Trim();
                    iter.SetCurrentLocation(endCell.End.Row, endCell.End.Column);
                }
            }
        }



        
        /// <summary>
        /// Finds the cell with the end header that corrisponds to the specified start header.
        /// </summary>
        /// <param name="worksheet">the worksheet that will soon be given formulas</param>
        /// <param name="startHeader">the cell containing a start header of a formula range</param>
        /// <returns>the cell with the end header, or null if no cell is found</returns>
        private static ExcelRange FindEndHeader(ExcelWorksheet worksheet, ExcelRange startHeader)
        {
            string targetText = "Total " + startHeader.Text;

            ExcelIterator iter = new ExcelIterator(worksheet, startHeader.Start.Row + 1, startHeader.Start.Column);

            return iter.GetCells(ExcelIterator.SHIFT_DOWN).FirstOrDefault(cell => cell.Text == targetText);
        }




        /// <summary>
        /// Gets all the string arguments required for the SummaryRowFormulaGenerator
        /// </summary>
        /// <param name="headers">the headers needed for the outer column of this report</param>
        /// <returns>a list of headers needed by the SummaryRowFormulaGenerator</returns>
        private static List<string> GetFinalSummaryHeaders(IEnumerable<string> headers)
        {
            StringBuilder firstFinalSummary = new StringBuilder("Total~");
            StringBuilder secondFinalSummary = new StringBuilder("Total:~");

            foreach (string header in headers)
            {
                string endHeader = header.Substring(header.IndexOf("=") + 1);

                firstFinalSummary.Append(endHeader);
                firstFinalSummary.Append(",");

                secondFinalSummary.Append(endHeader);
                secondFinalSummary.Append(",");
            }



            firstFinalSummary.Remove(firstFinalSummary.Length - 1, 1);
            secondFinalSummary.Remove(secondFinalSummary.Length - 1, 1);



            List<string> results = new List<string>(2);
            results.Add(firstFinalSummary.ToString());
            results.Add(secondFinalSummary.ToString());

            return results;
        }




        /// <summary>
        /// The AgedReceivables report sometimes has subtotals that need formulas and sometimes doesnt. This
        /// function checks if it has them or not.
        /// </summary>
        /// <param name="worksheet">the worksheet that might need the subtotals</param>
        /// <returns>true if the worksheet needs subtotals and false otherwise</returns>
        internal static bool AgedReceivablesNeedsSubtotals(ExcelWorksheet worksheet)
        {
            ExcelIterator iter = new ExcelIterator(worksheet);
            iter.GetFirstMatchingCell(cell => cell.Text.Trim() == "Description");
            int numSubtotals = iter.GetCells(ExcelIterator.SHIFT_DOWN)
                                    .Count(cell => cell.Text.Trim() == "Total");


            return numSubtotals >= 6;
        }
    }
}
