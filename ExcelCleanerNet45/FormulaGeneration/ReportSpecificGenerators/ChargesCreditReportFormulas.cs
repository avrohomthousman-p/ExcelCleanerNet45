using OfficeOpenXml;
using System;
using System.Collections.Generic;


namespace ExcelCleanerNet45.FormulaGeneration.ReportSpecificGenerators
{

    /// <summary>
    /// Implemenatation of IFormulaGenerator that searches for summary cells that also have
    /// header text in them (like "Total: $100), and splits them into seperate cells before
    /// adding a formula.
    /// </summary>
    internal class ChargesCreditReportFormulas : IFormulaGenerator
    {
        public IsDataCell dataCellDef = new IsDataCell(FormulaManager.IsDollarValue);



        public void InsertFormulas(ExcelWorksheet worksheet, string[] headers)
        {

            //stores the headers needed for the actual Formula Generator
            List<string> modifiedHeaders = new List<string>();
            ExcelIterator iter = new ExcelIterator(worksheet);

            foreach(string header in headers)
            {
                iter.SetCurrentLocation(1, 1);
                var matchingCells = iter.FindAllMatchingCells(cell => FormulaManager.TextMatches(cell.Text, header));

                foreach (ExcelRange cell in matchingCells)
                {
                    string modifiedHeader = SplitSummaryCell(worksheet, cell);
                    if(modifiedHeader != null)
                    {
                        modifiedHeaders.Add(modifiedHeader);
                    }
                }
                
            }


            //Now the actual generation of the formulas will be outsourced to the FullTableFormulaGenerator
            IFormulaGenerator actualGenerator = new FullTableFormulaGenerator();
            actualGenerator.SetDataCellDefenition(dataCellDef);
            actualGenerator.InsertFormulas(worksheet, modifiedHeaders.ToArray());
        }




        /// <summary>
        /// Splits the specified cell such that the data remains in the current cell
        /// and the text is moved to the cell before it.
        /// </summary>
        /// <param name="worksheet">the worksheet currently being given headers</param>
        /// <param name="cell">the cell we need to split</param>
        /// <returns>what the summary cell header is now as a result of our changes, or null if no change could be made</returns>
        private string SplitSummaryCell(ExcelWorksheet worksheet, ExcelRange cell)
        {
            int row = cell.Start.Row;
            int col = cell.End.Column;

            ExcelRange headerDestination = worksheet.Cells[row, col - 1];
            ExcelRange dataDestination = worksheet.Cells[row, col];



            if (!FormulaManager.IsEmptyCell(headerDestination))
            {
                return null;
            }


            //Now copy everything over
            int indexOfDollarSign = cell.Text.IndexOf('$');
            string header = cell.Text.Substring(0, indexOfDollarSign);
            string data = cell.Text.Substring(indexOfDollarSign);


            headerDestination.SetCellValue(0, 0, header);
            cell.CopyStyles(headerDestination);

            dataDestination.SetCellValue(0, 0, data);
            FormatAsDataCell(dataDestination);

            return header;
        }



        private void FormatAsDataCell(ExcelRange cell)
        {
            bool isNegative = cell.Text.StartsWith("(");

            cell.Style.Numberformat.Format = "$#,##0.00;($#,##0.00)";
            cell.Value = Double.Parse(StripNonDigits(cell.Text));

            if (isNegative)
            {
                cell.Value = (double)cell.Value * -1;
            }
        }



        /// <summary>
        /// Prepares text to be converted to a double by removing all commas, the preceding dollar sign, 
        /// and the surrounding parenthesis in the string if present.
        /// </summary>
        /// <param name="text">the text that should be cleaned</param>
        /// <returns>cleaned text that should be safe to parse to a double</returns>
        private string StripNonDigits(String text)
        {
            string replacementText;

            if (text.StartsWith("("))
            {
                replacementText = text.Substring(2, text.Length - 3);   //remove $, starting, and ending parenthesis
            }
            else
            {
                replacementText = text.Substring(1);                    //remove $ only
            }


            replacementText = replacementText.Replace(",", "");         //remove all commas


            return replacementText;
        }



        public void SetDataCellDefenition(IsDataCell isDataCell)
        {
            this.dataCellDef = isDataCell;
        }
    }
}
