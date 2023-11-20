using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace ExcelCleanerNet45.FormulaGeneration
{
    /// <summary>
    /// Implementation of IFormulaGenerator that adds formulas to the top of "sections" found inside data columns
    /// of the worksheet. A section is defined as a series of data cells that all corrispond to a single "key" which 
    /// appears on the top left of that section. As an example of this, look at the report VendorInvoiceReport.
    /// 
    /// This class is a lot like the PeriodicFormulaGenerator, but it puts the formulas on the top of each section
    /// instead of the bottom.
    /// 
    /// The first string in the list of arguments for this class should follow this pattern: r=[insert regex] 
    /// Where the regex should match the keys for each section. After that argument, the titles of each data column 
    /// that need formulas should be passed in as well (meaning, which columns need formulas at the top of each section).
    /// </summary>
    internal class PeriodicFormulasOnTop : PeriodicFormulaGenerator
    {
        //TODO

        protected override void ProcessFormulaRange(ExcelWorksheet worksheet, ref int row, int dataCol)
        {
            int start = row; //the formula range starts here, at the first non-empty cell


            //Now find the bottom of the formula range
            AdvanceToLastRow(worksheet, ref row);



            //Borrowing this function from the parent class, and using it to find the summary cell/top of
            //the NEXT section
            int lastRow = FindSummaryCellRow(worksheet, row, start, dataCol);
            if (lastRow == -1)
            {
                lastRow = worksheet.Dimension.End.Row;
            }




            //Insert formulas
            ExcelRange summaryCell = worksheet.Cells[start, dataCol];
            Console.WriteLine("adding formula to " + summaryCell.Address);

            string formula = FormulaManager.GenerateFormula(worksheet, start + 1, lastRow - 1, dataCol);

            FormulaManager.PutFormulaInCell(summaryCell, formula, false);
        }
    }
}
