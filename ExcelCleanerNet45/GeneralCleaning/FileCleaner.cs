using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Linq;
using System.Text.RegularExpressions;
using System.IO;

namespace ExcelCleanerNet45
{
    public class FileCleaner
    {


        /// <summary>
        /// Cleans an excel file
        /// </summary>
        /// <param name="sourceFile">the excel file in byte form</param>
        /// <param name="reportName">the file name of the original excel file</param>
        /// <param name="addFormulas">should be true if you also want formulas added to the report</param>
        /// <return>the excel file in byte form</return>
        public static byte[] OpenXLSX(byte[] sourceFile, string reportName, bool addFormulas=false)
        {

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;


            using (ExcelPackage package = new ExcelPackage(new MemoryStream(sourceFile)))
            {


                ExcelWorksheet worksheet;
                for (int i = 0; i < package.Workbook.Worksheets.Count; i++)
                {
                    worksheet = package.Workbook.Worksheets[i];

                    //If the worksheet is empty, Dimension will be null
                    if(worksheet.Dimension == null)
                    {
                        package.Workbook.Worksheets.Delete(i);
                        i--;
                        continue;
                    }

                    CleanWorksheet(worksheet, reportName);
                }


                byte[] results;
                if (addFormulas)
                {
                    results = FormulaManager.AddFormulas(package.GetAsByteArray(), reportName);
                }
                else
                {
                    results = package.GetAsByteArray();
                }
                
                

                Console.WriteLine("Workbook Cleanup complete");


                return results;

            }
        }



        /// <summary>
        /// Does the standard cleanup on the specified worksheet
        /// </summary>
        /// <param name="worksheet">the worksheet to be cleaned</param>
        /// <param name="reportName">the name of the report we are working on</param>
        public static void CleanWorksheet(ExcelWorksheet worksheet, string reportName)
        {

            DeleteHiddenRows(worksheet);


            RemoveAllHyperLinks(worksheet);


            RemoveAllMerges(worksheet, reportName);


            UnGroupAllRows(worksheet);


            CorrectCellDataTypes(worksheet);


            DoAdditionalCleanup(worksheet, reportName);

        }




        /// <summary>
        /// Deletes all hidden rows in the specified worksheet
        /// </summary>
        /// <param name="worksheet">the worksheet we are currently cleaning</param>
        private static void DeleteHiddenRows(ExcelWorksheet worksheet)
        {

            var end = worksheet.Dimension.End;

            for (int row = end.Row; row >= 1; row--)
            {
                if (worksheet.Row(row).Hidden == true)
                {
                    worksheet.DeleteRow(row);
                    Console.WriteLine("Deleted Hidden Row : " + row);
                }
                else if(RowIsSafeToDelete(worksheet, row))
                {
                    worksheet.DeleteRow(row);
                    Console.WriteLine("Deleted Very Small Row : " + row);
                }
            }
        }



        /// <summary>
        /// Checks if a row is empty and really really small and therefore no data would be lost if it was deleted
        /// </summary>
        /// <param name="worksheet">the worksheet currently being cleaned</param>
        /// <param name="rowNumber">the row being checked</param>
        /// <returns>true if the row is safe to delete becuase it has no data in it</returns>
        private static bool RowIsSafeToDelete(ExcelWorksheet worksheet, int rowNumber)
        {

            var row = worksheet.Row(rowNumber);
            if(row.Height >= 3)
            {
                return false;
            }



            //Check to see if the row is empty and can be deleted
            for (int colNumber = 1; colNumber <= worksheet.Dimension.Columns; colNumber++)
            {

                var cell = worksheet.Cells[rowNumber, colNumber];

                if(cell.Text != null && cell.Text.Length > 0) //if the cell has text in it (its not empty)
                {
                    return false; //unsafe to delete this row as it might have important text
                }
            }


            return true;

        }



        /// <summary>
        /// Removes all hyperlinks that are in any of the cells in the specified worksheet
        /// </summary>
        /// <param name="worksheet">the worksheet we are currently cleaning</param>
        private static void RemoveAllHyperLinks(ExcelWorksheet worksheet)
        {
            var start = worksheet.Dimension.Start;
            var end = worksheet.Dimension.End;

            for (int row = end.Row; row >= start.Row; row--)
            {
                for (int col = start.Column; col <= end.Column; ++col)
                {

                    var cell = worksheet.Cells[row, col];
                    StripCellOfHyperLink(cell, row, col);

                }
            }
        }



        /// <summary>
        /// Removes hyperlinks in the specified Excel Cell if any are present.
        /// </summary>
        /// <param name="cell">the cell whose hyperlinks should be removed</param>
        /// <param name="row">the row the cell is in</param>
        /// <param name="col">the column the cell is in</param>
        private static void StripCellOfHyperLink(ExcelRange cell, int row, int col)
        {
            if (cell.Hyperlink != null)
            {
                Console.WriteLine("Row=" + row.ToString() + " Col=" + col.ToString() + " Hyperlink=" + cell.Hyperlink);
                var val = cell.Value;
                cell.Hyperlink = null;
                cell.Value = val;
            }
        }




        /// <summary>
        /// Manages the unmerging
        /// </summary>
        /// <param name="worksheet">the worksheet whose cells must be unmerged</param>
        /// <param name="reportName">the name of the type of report being cleaned</param>
        private static void RemoveAllMerges(ExcelWorksheet worksheet, string reportName)
        {

            IMergeCleaner mergeCleaner = ReportMetaData.ChoosesCleanupSystem(reportName, worksheet.Index);

            try
            {
                mergeCleaner.Unmerge(worksheet);
            }
            catch(InvalidDataException e)
            {
                Console.WriteLine("Warning: Report " + reportName + " cannot be processed by the primary merge cleaner.");
                Console.WriteLine("Consider adding it to the list of reports that use the backup system.");
                Console.WriteLine("Error Message: " + e.Message);

                mergeCleaner = new BackupMergeCleaner();
                mergeCleaner.Unmerge(worksheet);
            }
            
        }



        /// <summary>
        /// Ungroups all grouped columns so that excel should not display a colapse or expand
        /// button (plus button or minus button) on the left margin.
        /// </summary>
        /// <param name="worksheet">the worksheet currently being cleaned</param>
        private static void UnGroupAllRows(ExcelWorksheet worksheet)
        {


            for (int row = 1; row <= worksheet.Dimension.Rows; row++)
            {
                var currentRow = worksheet.Row(row);

                if (currentRow.OutlineLevel > 0)
                {
                    currentRow.OutlineLevel = 0;
                }
            }
        }




        /// <summary>
        /// Checks all cells in the worksheet for data that is stored with bad formatting, and gives them proper formatting.
        /// Examples include dollar values that are being stored as text.
        /// </summary>
        /// <param name="worksheet">the worksheet currently being cleaned</param>
        private static void CorrectCellDataTypes(ExcelWorksheet worksheet)
        {
            for (int i = 1; i <= worksheet.Dimension.End.Row; i++)
            {
                for (int j = 1; j <= worksheet.Dimension.End.Column; j++)
                {

                    ExcelRange cell = worksheet.Cells[i, j];

                    

                    //Skip Empty Cells
                    if (cell.Text == null || cell.Text.Length == 0)
                    {
                        continue;
                    }


                    //there are some reports (like ReportPayablesRegister) with cells that should contain check 
                    //marks but instead have NaN in them (stored as a double). These need to be replaced with 
                    //actual check marks
                    if(cell.Value.GetType() == typeof(System.Double) && cell.Text == "NaN")
                    {
                        cell.Value = "ü";
                        continue;
                    }


                    //Skip Cells that already contain numbers
                    if (cell.Value.GetType() != typeof(string))
                    {
                        continue;
                    }





                    double unused;

                    // if it is a number but not a dollar value (like an ID), we want to store it as a number,
                    // and not get an excel warning
                    if (Double.TryParse(cell.Text, out unused))   
                    {

                        //Ignore the excel error that we have a number stored as a string
                        var error = worksheet.IgnoredErrors.Add(cell);
                        error.NumberStoredAsText = true;
                        continue;                                 //skip the formatting at the end of this if statement

                    }
                    // Some reports put just a dollar sign instead of a dollar amount of 0
                    else if(cell.Text == "$")
                    {
                        cell.Style.Numberformat.Format = "$#,##0.00;($#,##0.00)";
                        cell.Value = 0.0;
                    }
                    // if it a dollar value WITHOUT cents and stored as a string,  we want to convert it to an int and
                    // format it with a dollar sign and commas.
                    else if (IsDollarValue(cell.Text) && cell.Text.IndexOf('.') < 0)
                    {
                        bool isNegative = cell.Text.StartsWith("(");

                        cell.Style.Numberformat.Format = "$#,##0;($#,##0)";
                        cell.Value = Int32.Parse(CleanDollarValue(cell.Text)); //remove all non-digits, and parse to double

                        if (isNegative)
                        {
                            cell.Value = (int)cell.Value * -1;
                        }
                    }
                    // if it is a dollar value WITH cents stored as a string, we want to convert it to a double and format it
                    // with a dollar sign, commas, and 2 decimals
                    else if (IsDollarValue(cell.Text))
                    {

                        bool isNegative = cell.Text.StartsWith("(");

                        cell.Style.Numberformat.Format = "$#,##0.00;($#,##0.00)";
                        cell.Value = Double.Parse(CleanDollarValue(cell.Text)); //remove all non-digits, and parse to double

                        if (isNegative)
                        {
                            cell.Value = (double)cell.Value * -1;
                        }

                    }
                    //Replaces all dates formmated as mm/dd/yy with format mm/dd/yyyy
                    else if (IsDateWith2DigitYear(cell.Text))
                    {
                        string fourDigitYear = cell.Text.Substring(0, 6) + "20" + cell.Text.Substring(6);
                        cell.SetCellValue(0, 0, fourDigitYear);
                        continue;
                    }
                    //Percentages that are stored as text, should be stored as numbers with a % sign formatted in
                    else if (IsPercentage(cell.Text))
                    {
                        cell.Style.Numberformat.Format = "#0\\.00%;(#0\\.00%)";

                        bool isNegative = cell.Text.StartsWith("(");

                        string cleanedText = CleanPercentage(cell.Text);
                        double percentAsNumber = Double.Parse(cleanedText);

                        if (isNegative)
                        {
                            percentAsNumber *= -1;
                        }

                        cell.SetCellValue(0, 0, percentAsNumber);
                    }
                    else
                    {
                        continue; //If this data cannot be coverted to a number, skip the formatting below
                    }


                    
                    //When the alignment is set to general, text is left aligned but numbers are right aligned.
                    //Therefore if we change from text to number and we want to maintain alignment, we need to 
                    //change to right aligned.
                    if (cell.Style.HorizontalAlignment.Equals(ExcelHorizontalAlignment.General))
                    {
                        cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    }
                }
            }
        }




        /// <summary>
        /// Checks if the specified cell is a dollar value or not
        /// </summary>
        /// <param name="text">the text being checked</param>
        /// <returns>true if the specified text is a dollar value and false otherwise</returns>
        private static bool IsDollarValue(string text)
        {
            return text.StartsWith("$") || (text.StartsWith("($") && text.EndsWith(")"));
        }




        /// <summary>
        /// Prepares text to be converted to a double by removing all commas, the preceding dollar sign, 
        /// and the surrounding parenthesis in the string if present.
        /// </summary>
        /// <param name="text">the text that should be cleaned</param>
        /// <returns>cleaned text that should be safe to parse to a double</returns>
        private static string CleanDollarValue(String text)
        {
            string replacementText = RemoveParenthesis(text);


            replacementText = replacementText.Substring(1);             //Remove $


            replacementText = replacementText.Replace(",", "");         //remove all commas


            return replacementText;
        }



        /// <summary>
        /// Removes any parenthesis surrounding the text
        /// </summary>
        /// <param name="text">the text needing cleaning</param>
        /// <returns>the same text without the parenthesis</returns>
        private static string RemoveParenthesis(string text)
        {
            if(text.StartsWith("(") && text.EndsWith(")"))
            {
                return text.Substring(1, text.Length - 2);
            }
            else
            {
                return text;
            }
        }




        /// <summary>
        /// Checks if the specified text stores a date with a 2 digit year
        /// </summary>
        /// <param name="text">the text in question</param>
        /// <returns>true if the text matches the pattern of a date with a 2 digit year, and false otherwise</returns>
        private static bool IsDateWith2DigitYear(string text)
        {
            Regex reg = new Regex("^\\d\\d/\\d\\d/\\d\\d$");
            return reg.IsMatch(text);
        }




        /// <summary>
        /// Checks if the specified text is a percentage
        /// </summary>
        /// <param name="text">the text being checked</param>
        /// <returns>true if the text is a percentage stored in text, or false otherwise</returns>
        private static bool IsPercentage(string text)
        {
            return Regex.IsMatch(text, "^((100([.]00)?%)|([.]\\d\\d%)|(\\d{1,2}([.]\\d\\d)?%))$");
        }



        /// <summary>
        /// Removes all non digit characters from text so it can be converted into a double.
        /// </summary>
        /// <param name="text">the text that needs to be cleaned</param>
        /// <returns>a string that can be safely converted to a double</returns>
        private static string CleanPercentage(string text)
        {
            string cleanedText = RemoveParenthesis(text);

            cleanedText = cleanedText.Substring(0, cleanedText.Length - 1); //remove % sign

            return cleanedText;
        }



        /// <summary>
        /// Exceutes all report specific cleanup that needs to be done
        /// </summary>
        /// <param name="worksheet">the worksheet that is being cleaned</param>
        /// <param name="reportName">the report that is being cleaned</param>
        private static void DoAdditionalCleanup(ExcelWorksheet worksheet, string reportName)
        {
            if(ReportMetaData.NeedsSummaryCellsMoved(reportName, worksheet.Index))
            {
                MoveOutOfPlaceSummaryCells(worksheet);
            }
        }



        /// <summary>
        /// Moves all data cells in the last (rightmost) column over one cell to the left.
        /// This addresses a bug that leaves them one cell too far to the right.
        /// </summary>
        /// <param name="worksheet">the worksheet being cleaned</param>
        private static void MoveOutOfPlaceSummaryCells(ExcelWorksheet worksheet)
        {
            int col = worksheet.Dimension.End.Column;
            ExcelRange source, dest;
            
            for(int row = 1; row <= worksheet.Dimension.End.Row; row++)
            {
                source = worksheet.Cells[row, col];
                dest = worksheet.Cells[row, col - 1];
                if (CellShouldBeMoved(source, dest))
                {
                    source.CopyStyles(dest);
                    dest.Value = source.Value;
                    source.Value = null;
                }
                
            }
        }



        /// <summary>
        /// Checks if it is safe to transfer the data from the specified source cell to the specified destination.
        /// </summary>
        /// <param name="source">the source of the data</param>
        /// <param name="destination">the cell the data will be moved to</param>
        /// <returns>true if it is safe to do the transfer or false otherwise</returns>
        private static bool CellShouldBeMoved(ExcelRange source, ExcelRange destination)
        {
            if(!source.Text.StartsWith("$") && !source.Text.StartsWith("($"))
            {
                return false;
            }

            return destination.Text == null || destination.Text.Length == 0;
        }

    }

}
