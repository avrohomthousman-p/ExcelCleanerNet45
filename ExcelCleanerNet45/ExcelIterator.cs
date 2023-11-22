using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelCleanerNet45
{

    /// <summary>
    /// Helps iterate through an excel worksheet in any direction
    /// </summary>
    internal class ExcelIterator
    {
        public static readonly Tuple<int, int> SHIFT_UP = new Tuple<int, int>(-1, 0);
        public static readonly Tuple<int, int> SHIFT_DOWN = new Tuple<int, int>(1, 0);
        public static readonly Tuple<int, int> SHIFT_LEFT = new Tuple<int, int>(0, -1);
        public static readonly Tuple<int, int> SHIFT_RIGHT = new Tuple<int, int>(0, 1);



        private ExcelWorksheet worksheet;
        private int row;
        private int col;


        public ExcelIterator(ExcelWorksheet w)
        {
            this.worksheet = w;
            this.row = 1;
            this.col = 1;
        }



        public ExcelIterator(ExcelWorksheet w, int startRow, int startCol)
        {
            this.worksheet = w;
            this.row = startRow;
            this.col = startCol;
        }



        public ExcelIterator(ExcelIterator template)
        {
            this.worksheet = template.worksheet;
            this.row = template.row;
            this.col = template.col;
        }



        public int GetCurrentRow()
        {
            return row;
        }



        public int GetCurrentCol()
        {
            return col;
        }



        /// <summary>
        /// Gets the cell the iterator is currently referencing as a tuple with the row and column. The returned cell will
        /// be the same as the cell most recently returned by any of the iteration methods (GetCells or GetCellCoordinates)
        /// unless SetCurrentLocation was called since using those methods.
        /// </summary>
        /// <returns>the cell the iterator is refrencing</returns>
        public Tuple<int, int> GetCurrentLocation()
        {
            return new Tuple<int, int>(row, col);
        }



        /// <summary>
        /// Gets the cell the iterator is currently referencing. The returned cell will be the same as the cell most recently
        /// returned by any of the iteration methods (GetCells or GetCellCoordinates) unless SetCurrentLocation was called
        /// since using those methods.
        /// </summary>
        /// <returns>the cell the iterator is refrencing</returns>
        public ExcelRange GetCurrentCell()
        {
            return worksheet.Cells[row, col];
        }




        /// <summary>
        /// Sets the location the iterator is pointing to
        /// </summary>
        /// <param name="row">the row the iterator should point to</param>
        /// <param name="col">the column the iterator should point to</param>
        /// <exception cref="ArgumentOutOfRangeException">if the row or column are out of bounds</exception>
        public void SetCurrentLocation(int row, int col)
        {
            if (row < 1 || row > worksheet.Dimension.End.Row)
            {
                throw new ArgumentOutOfRangeException("Row " + row + " is out of range for this worksheet");
            }
            if (col < 1 || col > worksheet.Dimension.End.Column)
            {
                throw new ArgumentOutOfRangeException("Column " + col + " is out range for this worksheet");
            }



            this.row = row;
            this.col = col;
        }



        /// <summary>
        /// Advances the iterator to the next row.
        /// </summary>
        public void SkipRow()
        {
            this.row = Math.Min(this.row + 1, worksheet.Dimension.End.Row);
        }



        /// <summary>
        /// Advances the iterator to the next column.
        /// </summary>
        public void SkipColumn()
        {
            this.col = Math.Min(this.col + 1, worksheet.Dimension.End.Column);
        }



        /// <summary>
        /// Starting from the iterators current location, iterates and finds the first cell in the table that match 
        /// specified predicate.
        /// </summary>
        /// <param name="isDesiredCell">a predicate that returns true if this is the cell that you are looking for</param>
        /// <returns>the first cell found that matches the predicate, or null if no matching cell was found</returns>

        public ExcelRange GetFirstMatchingCell(Predicate<ExcelRange> isDesiredCell)
        {
            ExcelRange cell;


            for (; row <= worksheet.Dimension.End.Row; row++)
            {
                for (; col <= worksheet.Dimension.End.Column; col++)
                {
                    cell = worksheet.Cells[row, col];

                    if (isDesiredCell.Invoke(cell))
                    {
                        return cell;
                    }
                }

                
                col = 1;
            }


            return null;
        }



        /// <summary>
        /// Starting from the iterators current location, iterates and returns every cell in the table
        /// </summary>
        /// <returns>the row and column of each cell as a tuple</returns>
        public IEnumerable<Tuple<int, int>> FindAllCellCoordinates()
        {
            for (; row <= worksheet.Dimension.End.Row; row++)
            {
                for (; col <= worksheet.Dimension.End.Column; col++)
                {
                    yield return new Tuple<int, int>(row, col);
                }

                col = 1;
            }
        }




        /// <summary>
        /// Starting from the iterators current location, iterates and finds all cells in the whole table that match 
        /// specified predicate.
        /// </summary>
        /// <param name="isDesiredCell">a predicate that returns true if a cell should be returned</param>
        /// <returns>the row and column of the cell with matching the predicate as a tuple</returns>
        public IEnumerable<Tuple<int, int>> FindAllMatchingCoordinates(Predicate<ExcelRange> isDesiredCell)
        {
            ExcelRange cell;

            for (; row <= worksheet.Dimension.End.Row; row++)
            {
                for (; col <= worksheet.Dimension.End.Column; col++)
                {
                    cell = worksheet.Cells[row, col];

                    if (isDesiredCell.Invoke(cell))
                    {
                        yield return new Tuple<int, int>(row, col);
                    }
                }

                col = 1;
            }
        }




        /// <summary>
        /// Starting from the iterator's current location, iterates and returns each cell in the table.
        /// This method is an alternitive to FindAllMatchingCoordinates() that returns the cell
        /// itself instead of the coordinates.
        /// </summary>
        /// <returns>the ExcelRange object of each cell</returns>
        public IEnumerable<ExcelRange> FindAllCells()
        {
            foreach (Tuple<int, int> coordinates in FindAllCellCoordinates())
            {
                yield return worksheet.Cells[coordinates.Item1, coordinates.Item2];
            }
        }




        /// <summary>
        /// Starting from the iterator's current location, iterates and finds all cells in the whole table that match 
        /// specified predicate. This method is an alternitive to FindAllMatchingCoordinates(isDesiredCell) that 
        /// returns the cell itself instead of the coordinates.
        /// </summary>
        /// <param name="isDesiredCell">a predicate that returns true if a cell should be returned</param>
        /// <returns>the ExcelRange object of the cell with matching the predicate</returns>
        public IEnumerable<ExcelRange> FindAllMatchingCells(Predicate<ExcelRange> isDesiredCell)
        {
            foreach (Tuple<int, int> coordinates in FindAllMatchingCoordinates(isDesiredCell))
            {
                yield return worksheet.Cells[coordinates.Item1, coordinates.Item2];
            }
        }




        /// <summary>
        /// Iterates from the iterators current position, backwards, through every cell in the table.
        /// </summary>
        /// <returns>each cell the iterator passed through</returns>
        public IEnumerable<Tuple<int, int>> FindAllCellCoordinatesReverse()
        {
            for (; row > 0; row--)
            {
                for (; col > 0; col--)
                {
                    yield return new Tuple<int, int>(row, col);
                }

                col = worksheet.Dimension.End.Column;
            }
        }



        /// <summary>
        /// Iterates from the iterators current position, backwards, through every cell in the table.
        /// </summary>
        /// <returns>each cell the iterator passed through</returns>
        public IEnumerable<ExcelRange> FindAllCellsReverse()
        {
            foreach(Tuple<int, int> coordinate in FindAllCellCoordinatesReverse())
            {
                yield return worksheet.Cells[coordinate.Item1, coordinate.Item2];
            }
        }




        /// <summary>
        /// Iterates through every cell in the chosen direction until it reaches the end of the worksheet. After this
        /// operation the iterator will reference the last element in the worksheet (in the direction of the iteration).
        /// </summary>
        /// <param name="shift">the direction to move in, represented as a tuple with row change, and column change</param>
        /// <returns>the cells it iterates through</returns>
        public IEnumerable<ExcelRange> GetCells(Tuple<int, int> shift)
        {
            while (!OutOfBounds())
            {
                ExcelRange cell = worksheet.Cells[row, col];

                yield return cell;

                row += shift.Item1;
                col += shift.Item2;
            }


            //Undo final increment so the iterator should point to the last cell in the worksheet
            row -= shift.Item1;
            col -= shift.Item2;
        }



        /// <summary>
        /// Iterates through every cell in the chosen direction until it reaches the end of the worksheet or the 
        /// specified predicate evaluates to true. After this operation the iterator will reference the cell that
        /// made our predicate true or the last cell before the iterator would go out of bounds.
        /// </summary>
        /// <param name="shift">the direction to move in, represented as a tuple with row change, and column change</param>
        /// <param name="stopIf">the condition that signal we should stop iterating</param>
        /// <returns>the cells it iterates through</returns>
        public IEnumerable<ExcelRange> GetCells(Tuple<int, int> shift, Predicate<ExcelRange> stopIf)
        {
            while (!OutOfBounds())
            {
                ExcelRange cell = worksheet.Cells[row, col];

                if (stopIf.Invoke(cell))
                {
                    break;
                }

                
                yield return cell;

                row += shift.Item1;
                col += shift.Item2;
            }


            //Undo final increment so that the iterator points to the last cell before we stopped
            row -= shift.Item1;
            col -= shift.Item2;
        }





        /// <summary>
        /// Iterates through every cell in the chosen direction until it reaches the end of the worksheet. The only difference
        /// between this method and the GetCells(shift) method is that this method returns the row and column of each cell, not
        /// the ExcelRange object.
        /// </summary>
        /// <param name="shift">the direction to move in, represented as a tuple with row change, and column change</param>
        /// <returns>the row and column of the cells it iterates through (stored in a tuple)</returns>
        public IEnumerable<Tuple<int, int>> GetCellCoordinates(Tuple<int, int> shift)
        {
            foreach(ExcelRange cell in GetCells(shift))
            {
                yield return new Tuple<int, int>(row, col);
            }
        }




        /// <summary>
        /// Iterates through every cell in the chosen direction until it reaches the end of the worksheet of the 
        /// specified predicate evaluates to true. The only difference between this method and the 
        /// GetCells(shift, stopIf) method is that this method returns the row and column of each cell, not
        /// the ExcelRange object.
        /// </summary>
        /// <param name="shift">the direction to move in, represented as a tuple with row change, and column change</param>
        /// <returns>the row and column of the cells it iterates through (stored in a tuple)</returns>
        public IEnumerable<Tuple<int, int>> GetCellCoordinates(Tuple<int, int> shift, Predicate<ExcelRange> stopIf)
        {
            foreach (ExcelRange cell in GetCells(shift, stopIf))
            {
                yield return new Tuple<int, int>(row, col);
            }
        }




        /// <summary>
        /// Continues iteration from current cell, and keeps moving untill the specified condition returns false.
        /// After this operation, the iterator will be referencing the element that made the condition false, or
        /// the last element in the worksheet if we went out of bounds.
        /// </summary>
        /// <param name="shift">the direction to iterate in, stored as a tuple of (row increment, col increment)</param>
        /// <param name="condition">the condition that tells the iterator to keep going</param>
        public void SkipWhile(Tuple<int, int> shift, Predicate<ExcelRange> condition)
        {

            bool keepGoing = true;
            ExcelRange cell;

            while(!OutOfBounds() && keepGoing)
            {
                cell = worksheet.Cells[row, col];
                keepGoing = condition.Invoke(cell);
                row += shift.Item1;
                col += shift.Item2;
            }



            //undo final increment so we reference the element that made the predicate false
            //or the last element in the worksheet
            row -= shift.Item1;
            col -= shift.Item2;
        }




        /// <summary>
        /// Convienece method to skip all empty cells and continue iteration after them. This method is
        /// equivilent to calling SkipWhile with a predicate that checks if the cell is empty.
        /// </summary>
        /// <param name="shift">the direction to iterate in</param>
        public void SkipEmptyCells(Tuple<int, int> shift)
        {
            SkipWhile(shift, cell => FormulaManager.IsEmptyCell(cell));
        }



        /// <summary>
        /// Checks if the iterator is out of bounds of the worksheet
        /// </summary>
        /// <returns>true if the iterator has gone out of bounds, and false otherwise</returns>
        private bool OutOfBounds()
        {
            return col < 1 || col > worksheet.Dimension.End.Column || row < 1 || row > worksheet.Dimension.End.Row;
        }
    }
}
