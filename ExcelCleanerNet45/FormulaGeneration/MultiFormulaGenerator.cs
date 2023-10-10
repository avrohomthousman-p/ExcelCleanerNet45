using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelCleanerNet45.FormulaGeneration
{
    /// <summary>
    /// Implementation of IFormulaGenerator interface that calls two other formula generators to do the work. 
    /// This is a convienent tool that lets you add formuls in two (or three) different ways without making a new 
    /// implementation for it. When passing headers to this class, all headers intended for the first formula generator should
    /// start with a 1, and those intended for the second should start with a 2, and so on for the third (which is optional).
    /// </summary>
    internal class MultiFormulaGenerator : IFormulaGenerator
    {
        private IFormulaGenerator firstGenerator, secondGenerator, thirdGenerator;
        private IsDataCell dataCellDef = null; //use the default for each formula generator


        public MultiFormulaGenerator(IFormulaGenerator first, IFormulaGenerator second, IFormulaGenerator third = null)
        {
            this.firstGenerator = first;
            this.secondGenerator = second;
            this.thirdGenerator = third;
        }




        public void InsertFormulas(ExcelWorksheet worksheet, string[] headers)
        {
            //Seperate arguments for the first and second formula generator and remove the leading digit
            string[] argumentsForFirst = headers.Where(text => text.StartsWith("1")).Select(text => text.Substring(1)).ToArray();
            string[] argumentsForSecond = headers.Where(text => text.StartsWith("2")).Select(text => text.Substring(1)).ToArray();
            string[] argumentsForThird = headers.Where(text => text.StartsWith("3")).Select(text => text.Substring(1)).ToArray();

            firstGenerator.InsertFormulas(worksheet, argumentsForFirst);
            secondGenerator.InsertFormulas(worksheet, argumentsForSecond);

            if(thirdGenerator != null)
            {
                thirdGenerator.InsertFormulas(worksheet, argumentsForThird);
            }
        }




        public void SetDataCellDefenition(IsDataCell isDataCell)
        {
            this.dataCellDef = isDataCell;
            firstGenerator.SetDataCellDefenition(isDataCell);
            secondGenerator.SetDataCellDefenition(isDataCell);

            if(thirdGenerator != null)
            {
                thirdGenerator.SetDataCellDefenition(isDataCell);
            }
        }
    }
}
