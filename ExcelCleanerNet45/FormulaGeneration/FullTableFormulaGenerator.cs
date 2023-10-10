using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelCleanerNet45
{

    //Used to tell the Formula Generator when to stop adding cells to a formula range
    public delegate bool IsBeyondFormulaRange(ExcelRange cell);



    /// <summary>
    /// Implementation of the IFormulaGenerator interface that searches for a row with the specifed header
    /// and adds a formula to each data cell in that row that spans as far up as it can.
    /// </summary>
    internal class FullTableFormulaGenerator : IFormulaGenerator
    {

        private ExcelIterator iter;
        protected IsBeyondFormulaRange beyondFormulaRange;
        protected IsDataCell isDataCell;



        /// <summary>
        /// Constructs a FullTableFormulaGenerator that considers a forumal range to end when it encounters a 
        /// cell that is either empty or is not a data cell (does not start with a $ sign). If you want to 
        /// change this condition, pass in your own predicate to this constructor.
        /// </summary>
        public FullTableFormulaGenerator()
        {
            //Set default implemenatations of all delegates
            isDataCell = new IsDataCell(FormulaManager.IsDollarValue);
            beyondFormulaRange = new IsBeyondFormulaRange(IsEmptyOrNonDataCell);
        }



        /// <inheritdoc/>
        public void SetDataCellDefenition(IsDataCell isDataCell)
        {
            this.isDataCell = isDataCell;
        }




        /// <summary>
        /// Sets the implementation for how this object determans if a cell is outside a formula range.
        /// Note, two implementations are predefined and can be paased into this method: IsNonDataCell and
        /// IsEmptyOrNonDataCell
        /// </summary>
        /// <param name="altImplemenation">the implementation that should be used</param>
        public void SetDefenitionForBeyondFormulaRange(IsBeyondFormulaRange altImplemenation)
        {
            this.beyondFormulaRange = altImplemenation;
        }



        /// <inheritdoc/>
        public virtual void InsertFormulas(ExcelWorksheet worksheet, string[] headers)
        {
            iter = new ExcelIterator(worksheet);

            //for each header in the report that needs a formula 
            foreach (string header in headers)              
            {

                //Ensure that the header was intended for this class and not the DistantRowsFormulaGenerator class
                if (FormulaManager.IsNonContiguousFormulaRange(header))
                {
                    continue;
                }


                iter.SetCurrentLocation(1, 1);
                var allHeaderCoordinates = iter.FindAllMatchingCoordinates(cell => FormulaManager.TextMatches(cell.Text, header));

                
                //Find each instance of that header and add formulas
                foreach(var coordinates in allHeaderCoordinates)
                {
                    FillInFormulas(worksheet, coordinates.Item1, coordinates.Item2);
                }
            }

        }




        /// <summary>
        /// Inserts the formulas in each cell in the formula range that requires it.
        /// </summary>
        /// <param name="worksheet">the worksheet currently being given formulas</param>
        /// <param name="row">the row of the header</param>
        /// <param name="col">the column of the header</param>
        protected virtual void FillInFormulas(ExcelWorksheet worksheet, int row, int col)
        {
            iter.SetCurrentLocation(row, col);

            foreach (ExcelRange cell in iter.GetCells(ExcelIterator.SHIFT_RIGHT))
            {
                if (FormulaManager.IsEmptyCell(cell) || !isDataCell(cell))
                {
                    //Console.WriteLine($"cell {cell.Address} was skipped");
                    continue;
                }


                int topRowOfRange = FindTopRowOfFormulaRange(worksheet, row, col);

                cell.Formula = FormulaManager.GenerateFormula(worksheet, topRowOfRange, row - 1, iter.GetCurrentCol());
                cell.Style.Locked = true;

                Console.WriteLine("Cell " + cell.Address + " has been given this formula: " + cell.Formula);
            }

        }




        /// <summary>
        /// Given the coordinates to the bottom cell in a formula range, checks how far up the range goes
        /// </summary>
        /// <param name="worksheet">the worksheet in need of formulas</param>
        /// <param name="row">the row number of the bottom cell in the range</param>
        /// <param name="col">the column number of the bottom cell in the range</param>
        /// <returns>the row number of the top most cell thats still part of the formula range</returns>
        protected virtual int FindTopRowOfFormulaRange(ExcelWorksheet worksheet, int row, int col)
        {
            ExcelIterator iterateOverFormulaRange = new ExcelIterator(iter);

            Tuple<int, int> cellAboveRange = iterateOverFormulaRange
                .GetCellCoordinates(ExcelIterator.SHIFT_UP, stopIf: new Predicate<ExcelRange>(beyondFormulaRange))
                .Last();


            return cellAboveRange.Item1;

        }





        //Just convienences that you can pass to this classes setter methods


        /// <summary>
        /// Checks if the specified cell is not part of the formula range becuase it is an empty cell or 
        /// it contains text and not data. This method is just a convienence that you can pass to this classes'
        /// SetDefenitionForBeyondFormulaRange method if that is the behavior you want.
        /// </summary>
        /// <param name="cell">the cell being checked</param>
        /// <returns>true if the cell is empty or if it contains text that isnt data, or false otherwise</returns>
        public virtual bool IsEmptyOrNonDataCell(ExcelRange cell)
        {
            return FormulaManager.IsEmptyCell(cell) || !isDataCell(cell);
        }




        /// <summary>
        /// Checks if the specified cell is not part of the formula range becuase it contains text that is not data.
        /// This method is just a convienence that you can pass to this classes'
        /// SetDefenitionForBeyondFormulaRange method if that is the behavior you want.
        /// </summary>
        /// <param name="cell">the cell being checked</param>
        /// <returns>true if the cell contains text that isnt data (and isnt empty), or false otherwise</returns>
        public virtual bool IsNonDataCell(ExcelRange cell)
        {
            return !FormulaManager.IsEmptyCell(cell) && !isDataCell(cell);
        }
            
    }
}
