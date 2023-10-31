using OfficeOpenXml;
using System;
using System.IO;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using OfficeOpenXml.FormulaParsing;
using ExcelCleanerNet45.FormulaGeneration;

namespace ExcelCleanerNet45
{
    /// <summary>
    /// Replaces static values in excel files with formulas that will change when the data is updated
    /// </summary>
    public class FormulaManager
    {



        /// <summary>
        /// Adds all necissary formulas to the appropriate cells in the specified file
        /// </summary>
        /// <param name="sourceFile">the excel file needing formulas, stored as an array/stream of bytes</param>
        /// <param name="reportName">the name of the report</param>
        /// <returns>the byte stream/arrray of the modified file</returns>
        public static byte[] AddFormulas(byte[] sourceFile, string reportName)
        {
            
            using (ExcelPackage package = new ExcelPackage(new MemoryStream(sourceFile)))
            {

                string[] headers;
                ExcelWorksheet worksheet;

                for (int i = 0; i < package.Workbook.Worksheets.Count; i++)
                {
                    worksheet = package.Workbook.Worksheets[i];


                    //If the worksheet is empty, Dimension will be null
                    if (worksheet.Dimension == null)
                    {
                        package.Workbook.Worksheets.Delete(i);
                        i--;
                        continue;
                    }


                    //Get the formula generator object that will insert the formulas
                    IFormulaGenerator formulaGenerator = ReportMetaData.ChooseFormulaGenerator(reportName, i);



                    if(formulaGenerator == null)    //if this worksheet doesnt need formulas
                    {
                        continue;                   //skip this worksheet
                    }



                    //get the arguments that are required for the formula generator
                    headers = ReportMetaData.GetFormulaGenerationArguments(reportName, i);


                    //Actually add the formulas
                    formulaGenerator.InsertFormulas(worksheet, headers);


                    //Many reports require some additional formulas that will be added by the SummaryRowFormulaGenerator
                    SummaryRowFormulaGenerator summaryGenerator = new SummaryRowFormulaGenerator();
                    summaryGenerator.InsertFormulas(worksheet, headers);
                }


                return package.GetAsByteArray();
            }

        }





        /* Some utility methods needed by the Formula generators */


        /// <summary>
        /// Checks if a cell is empty (has no text)
        /// </summary>
        /// <param name="cell">the cell being checked</param>
        /// <returns>true if the cell has no text and false otherwise</returns>
        internal static bool IsEmptyCell(ExcelRange cell)
        {
            return (cell.Text == null || cell.Text.Length == 0);
        }




        /// <summary>
        /// Checks if the specified cell contains a percentage in it
        /// </summary>
        /// <param name="cell">the cell being checked</param>
        /// <returns>true if the cell contains a percentage, and false otherwise</returns>
        internal static bool IsPercentage(ExcelRange cell)
        {
            return TextMatches(cell.Text, "(100([.]00)?%)|([.]\\d\\d%)|(\\d{1,2}([.]\\d\\d)?%)"); //"1?\\d\\d(\\.\\d\\d)?%"
        }



        /// <summary>
        /// Checks if a cell contains a dollar value. This is used as a default implementation for IsDataCell in 
        /// formula managers. Formula managers can be set to use a different defenition for a data cell 
        /// by calling the method IFormulaManager.SetDataCellDefenition(  specify alternate implementation  )
        /// </summary>
        /// <param name="cell">the cell being checked</param>
        /// <returns>true if the cell contains a dollar value and false otherwise</returns>
        internal static bool IsDollarValue(ExcelRange cell)
        {
            return cell.Text.StartsWith("$") || (cell.Text.StartsWith("($") && cell.Text.EndsWith(")"));
        }




        /// <summary>
        /// Checks if the contents of a cell is a integer
        /// </summary>
        /// <param name="cell">the cell being checked</param>
        /// <returns>true if the cell contains an integer (and nothing else) and false otherwise</returns>
        internal static bool IsIntegerValue(ExcelRange cell)
        {
            return TextMatches(cell.Text, "0|(-?[1-9]\\d*)");
        }




        /// <summary>
        /// Checks if the contents of a cell is a integer whose digits are seperated by commas (e.g. 1,000)
        /// </summary>
        /// <param name="cell">the cell being checked</param>
        /// <returns>true if the cell contains an integer (and nothing else) and false otherwise</returns>
        internal static bool IsIntegerWithCommas(ExcelRange cell)
        {
            return TextMatches(cell.Text, "0|(-?[1-9]\\d{0,2}(,\\d{3})*)");
        }




        /// <summary>
        /// Checks if the specified cell has a formula in it
        /// </summary>
        /// <param name="cell">the cell being checked</param>
        /// <returns>true if the cell has a formula in it and false otherwise</returns>
        internal static bool CellHasFormula(ExcelRange cell)
        {
            return cell.Formula != null && cell.Formula.Length > 0;
        }




        /// <summary>
        /// Generates the formula for the cells in the given range. Note: the range should only include the 
        /// cells that are to be included in the formula. Not the that cell that will contain the formula itself
        /// or any cells above the range.
        /// </summary>
        /// <param name="worksheet">the worksheet currently getting formulas</param>
        /// <param name="startRow">the first data cell to be included in the formula</param>
        /// <param name="endRow">the last data cell to be included in the formula</param>
        /// <param name="col">the column the formula is for</param>
        /// <returns>the proper formula for the specified formula range</returns>
        internal static string GenerateFormula(ExcelWorksheet worksheet, int startRow, int endRow, int col)
        {
            ExcelRange cells = worksheet.Cells[startRow, col, endRow, col];

            return "SUM(" + cells.Address + ")";
        }




        /// <summary>
        /// Inserts the specified formula into the cell, and performs any addiotional operations accosiated with that task
        /// </summary>
        /// <param name="cell">the cell recieving the formula</param>
        /// <param name="formula">the formula to be inserted in the cell</param>
        /// <param name="output">optional argument. If true, a console log will be printed displaying the formula added</param>
        internal static void PutFormulaInCell(ExcelRange cell, string formula, bool output = true)
        {
            cell.Formula = formula;
            cell.Style.Locked = true;
            cell.Style.Hidden = false;

            //This causes the results of each formula to be cached in the file, so it will be visible
            //when the file is opened in protected mode.
            cell.Calculate();


            if (output)
            {
                Console.WriteLine("Cell " + cell.Address + " has been given this formula: " + cell.Formula);
            }
        }




        /// <summary>
        /// Checks if a header is intened for the DistantRowsFormulaGenerator or not.
        /// </summary>
        /// <param name="header">the header in question</param>
        /// <returns>true if the specified header is intended for the DistantRowsFormulaGenerator class, and false otherwise</returns>
        internal static bool IsNonContiguousFormulaRange(string header)
        {
            return header.IndexOf('~') >= 0;
        }


        

        /// <summary>
        /// Checks if the specified text matches (in its entirety) the specified regex.
        /// </summary>
        /// <param name="text">the text to be matched</param>
        /// <param name="pattern">the pattern the text should match</param>
        /// <returns>true if the text matches the pattern and false otherwise</returns>
        internal static bool TextMatches(string text, string pattern)
        {
            return Regex.IsMatch(text.Trim(), "^" + pattern + "$");
        }

    }


}
