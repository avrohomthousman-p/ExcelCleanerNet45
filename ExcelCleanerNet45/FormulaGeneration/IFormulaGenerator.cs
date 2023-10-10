using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelCleanerNet45
{
    public delegate bool IsDataCell(ExcelRange cell);


    /// <summary>
    /// A class implementing this interface should have the ability to add formulas to a worksheet.
    /// 
    /// An implementing class does not necessarily add all the formulas that the worksheet needs. Often
    /// more than one implementation of this interface will need to be used to add all the formulas needed.
    /// 
    /// Formula Generators are told where to add the formulas based on the "headers", an array of strings passed to the
    /// InsertFormulas method below. Each string in the array is a regex that matches the text of at least one of 
    /// the cells in the worksheet. That cell is the "header cell" and is where the formula should be added.
    /// 
    /// Note: the exact details of how the header cell aligns with the formula cell (the cell that gets the summary formula) 
    /// depends on the implementation. For most formula generators, the formula cells will be on the same row
    /// as the header cell, somewhere to the right, and the formula range for each formula cell will be above that
    /// formula cell.
    /// 
    /// For more precise details on how each formula generator finds its formula cells and how the headers should be
    /// written, check the documentation for that formula generator.
    /// </summary>
    internal interface IFormulaGenerator
    {
        /// <summary>
        /// Finds each instance of each header in the worksheet and gives its corisponding columns a formula.
        /// The term "header" here refers to minor headers in the worksheet with text like "Total Income" that 
        /// indicate that the row they are in is a summary row whose value should be calculated by a formula and 
        /// not just be a static value.
        /// </summary>
        /// <param name="worksheet">the worksheet in need of formulas</param>
        /// <param name="headers">an array of all minor headers that are right in front of cells requiring formulas</param>
        void InsertFormulas(ExcelWorksheet worksheet, string[] headers);



        /// <summary>
        /// Changes how the formula generator defines a data cell.
        /// </summary>
        /// <param name="isDataCell">a function to use to check if a cell is a data cell</param>
        void SetDataCellDefenition(IsDataCell isDataCell);
    }
}
