using ExcelCleanerNet45.FormulaGeneration;
using ExcelCleanerNet45.FormulaGeneration.ReportSpecificGenerators;
using ExcelCleanerNet45.GeneralCleaning;
using ExcelCleanerNet45;
using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography.Xml;
using System.Text.RegularExpressions;
using OfficeOpenXml;
using System.Linq;

namespace ExcelCleanerNet45
{

    /// <summary>
    /// Stores and makes accesible all meta data about reports, like what merge cleaner and formula generator to
    /// use.
    /// </summary>
    internal static class ReportMetaData
    {

        private static readonly string anyMonth = "(January|February|March|April|May|June|July|August|September|October|November|December)";
        private static readonly string anyDate = "\\d{1,2}/\\d{1,2}/\\d{4}";
        private static readonly string anyYear = "[12]\\d\\d\\d";
        private static readonly string anyProperty = "([A-Z][a-z]+)( (([A-Z][a-z]+)|(at)))*";




        /// <summary>
        /// Factory method for choosing a version of the merge cleanup code that would work best for the specified report
        /// </summary>
        /// <param name="reportName">the name of the report that needs unmerging</param>
        /// <param name="reportVersion">teh version of the report being unmerged. Null or empty if only one version exists</param>
        /// <param name="worksheetNumber">the worksheet withing the report that needs unmerging</param>
        /// <returns>an instance of IMergeCleaner that should be used to clean the specified worksheet</returns>
        internal static IMergeCleaner ChoosesCleanupSystem(string reportName, string reportVersion, int worksheetNumber)
        {
            switch (reportName)
            {
                case "SummaryReport":
                case "TrialBalance":
                case "TrialBalanceVariance":
                case "CashFlow":
                case "UnitInfoReport":
                case "ReportCashReceiptsSummary":
                    return new BackupMergeCleaner();




                case "ReportTenantSummary":
                    AbstractMergeCleaner m = new BackupMergeCleaner();    
                    m.AddCleanupJob(AdditionalCleanupJobs.SetColumnWidthsForReportTenantSummary);
                    return m;



                case "VendorPropertyReport":
                    m = new BackupMergeCleaner();

                    //ensure each column isnt small enough to hide any data
                    m.AddCleanupJob(worksheet => {
                        AdditionalCleanupJobs.SetColumnsToMinimumWidth(worksheet, 1.5);
                    });


                    return m;



                case "BalanceSheetDrillthrough":
                    m = new BackupMergeCleaner();

                    //Make sure all data cells in the data column have the correct alignment
                    m.AddCleanupJob(AdditionalCleanupJobs.RealignDataColumn);

                    return m;



                case "ReportAccountBalances":
                    m = new BackupMergeCleaner();

                    //this report has some of its data rows frozen (they stick to the screen as you scroll away)
                    //we dont want that
                    m.AddCleanupJob(worksheet => {
                        //First find the first data cell, so we can freeze only the rows above it
                        ExcelIterator iter = new ExcelIterator(worksheet);
                        int row = iter.GetFirstMatchingCell(cell => FormulaManager.IsDollarValue(cell)).Start.Row;

                        worksheet.View.UnFreezePanes();
                        worksheet.View.FreezePanes(row, worksheet.Dimension.End.Column);
                    });

                    return m;



                case "ProfitAndLossStatementDrillthrough":
                case "ProfitAndLossStatementDrillThrough":
                    m = new BackupMergeCleaner();
                    m.MoveMajorHeaders = false;
                    return m;



                case "ProfitAndLossStatementByJob":
                    m = new PrimaryMergeCleaner();
                    m.MoveMajorHeaders = false;
                    m.AddCleanupJob(worksheet => AdditionalCleanupJobs.ApplyColumnMaxWidth(worksheet, 18));
                    return m;



                case "ProfitAndLossStatementByPeriod":
                case "ProfitAndLossComp":
                    m = new PrimaryMergeCleaner();

                    //Header with text Cash Basis is often not aligned correctly
                    m.AddCleanupJob(worksheet => 
                    {
                        AdditionalCleanupJobs.RealignSingleHeader(worksheet, "Cash Basis",
                                                OfficeOpenXml.Style.ExcelHorizontalAlignment.Left);
                    });


                    if(reportName == "ProfitAndLossComp")
                    {
                        m.AddCleanupJob(worksheet => AdditionalCleanupJobs.ApplyColumnMaxWidth(worksheet, 16));
                    }


                    return m;



                case "InvoiceDetail":
                    m = new BackupMergeCleaner();

                    //Header with text Cash Basis is often not aligned correctly
                    m.AddCleanupJob(worksheet =>
                    {
                        AdditionalCleanupJobs.RealignSingleHeader(worksheet, "Payments",
                                                OfficeOpenXml.Style.ExcelHorizontalAlignment.Left);
                    });


                    return m;




                case "ProfitAndLossExtendedVariance":
                    m = new ExtendedVarianceCleaner();
                    m.MoveMajorHeaders = false;
                    return m;



                case "ReportOutstandingBalance":
                    switch (worksheetNumber)
                    {
                        case 1:
                            return new BackupMergeCleaner();
                        default:
                            return new PrimaryMergeCleaner();
                    }



                case "PaymentsHistory":
                    switch (worksheetNumber)
                    {
                        case 1:
                            return new ReAlignDataCells("Grand Total:");

                        default:
                            return new PrimaryMergeCleaner();
                    }



                case "RentRollHistory":
                    switch (worksheetNumber)
                    {
                        case 1:
                            m = new ReAlignDataCells("Vacancy %");

                            //the data in the first column gets hidden as the row is collapsed
                            m.AddCleanupJob(new Action<ExcelWorksheet>(worksheet =>
                            {
                                worksheet.Column(1).Width = 1.5;
                            }));


                            return m;

                        default:
                            return new PrimaryMergeCleaner();
                    }



                case "SecurityReport":
                    m = new PrimaryMergeCleaner();
                    m.AddCleanupJob(worksheet => 
                    {
                        //resize the last row
                        int lastRow = worksheet.Dimension.End.Row;
                        worksheet.Row(lastRow).Height = 16.0;
                    });

                    return m;


                    
                case "VendorInvoiceReportWithJournalAccounts":
                    m = new PrimaryMergeCleaner();

                    m.AddCleanupJob(worksheet => {
                        AdditionalCleanupJobs.SetColumnToWrapText(worksheet, "Journal Account");
                    });

                    return m;




                case "VendorInvoiceReport":
                    m = new PrimaryMergeCleaner();

                    m.AddCleanupJob(worksheet => {
                        ExcelRange topOfColumn = AdditionalCleanupJobs.GetCellWithText(worksheet, "Description");
                        AdditionalCleanupJobs.SetColumnToWrapText(worksheet, topOfColumn.End.Row + 1, topOfColumn.Start.Column);
                        AdditionalCleanupJobs.ResizeColumnAndDeleteTheNext2(worksheet, topOfColumn.Start.Column);
                    });

                    return m;




                case "DistributionsReport":
                    m = new PrimaryMergeCleaner();
                    m.AddCleanupJob(worksheet => {
                        AdditionalCleanupJobs.DeleteEmptyCellsForDistributionsReport(worksheet, "Distribution", "\\d+ Distribution(s)?");
                    });

                    return m;



                case "PropertyReport":
                    m = new PrimaryMergeCleaner();

                    m.AddCleanupJob(AdditionalCleanupJobs.SetAllColumnsToWrapText);

                    return m;



                case "Budget":
                    return new ReAlignDataCells();


                case "UnitInvoiceReport":
                    m = new PrimaryMergeCleaner();

                    switch (reportVersion)
                    {
                        case "JustAmounts":
                            m.AddCleanupJob(AdditionalCleanupJobs.LeftAlignAllHeadersInFirstRow);
                            break;



                        default:
                            m.AddCleanupJob(worksheet => {
                                AdditionalCleanupJobs.SetColumnToWrapText(worksheet, "Description");
                                AdditionalCleanupJobs.SetColumnToMinimumWidth(worksheet, "Due Date", 10.25);
                            });
                            break;
                    }


                    return m;



                case "InvoiceList":
                    m = new PrimaryMergeCleaner();
                    m.AddCleanupJob(AdditionalCleanupJobs.FixSummariesOfInvoiceList);
                    return m;



                case "ProfitAndLossBudget":
                    return new ReAlignMergeCells();


                case "BankReconcilliation":
                    return new SetDefaultColumnWidth();


                default:
                    return new PrimaryMergeCleaner();
            }
        }



        /// <summary>
        /// Checks if the specified worksheet needs to have its summary cells shifted one cell to the left.
        /// Due to a bug in the report generator, some reports have their summary cells one cell too far
        /// to the right.
        /// </summary>
        /// <param name="reportName">the name of the report the worksheet is from</param>
        /// <param name="reportVersion">teh version of the report being unmerged. Null or empty if only one version exists</param>
        /// <param name="worksheetIndex">the zero based index of the worksheet</param>
        /// <returns>true if the worksheets needs its summary cells moved, and false otherwise</returns>
        internal static bool NeedsSummaryCellsMoved(string reportName, string reportVersion, int worksheetIndex)
        {
            switch (reportName)
            {
                //TODO: add the other reports with this issue
                case "ReportOutstandingBalance":
                //case "ProfitAndLossBudget":
                    return true;


                default:
                    return false;
            }
        }




        /// <summary>
        /// Factory method for choosing the implementation of the IFormulaGenerator interface that should be used to add formulas
        /// to the specified report.
        /// </summary>
        /// <param name="reportName">the name of the report that needs formulas</param>
        /// <param name="worksheetNum">the index of the worksheet that needs formulas</param>
        /// <param name="workbook">the full workbook we are in. Sometimes needed to check on how many worksheets there are</param>
        /// <param name="reportVersion">teh version of the report being unmerged. Null or empty if only one version exists</param>
        /// <returns>
        /// an implemenation of the IFormulaGenerator interface that should be used to add the formulas,
        /// or null if the worksheet doesnt need formulas
        /// </returns>
        internal static IFormulaGenerator ChooseFormulaGenerator(string reportName, int worksheetNum, ExcelWorkbook workbook, string reportVersion)
        {

            FullTableFormulaGenerator formulaGenerator;


            switch (reportName)
            {
                case "ProfitAndLossStatementDrillthrough":
                case "ProfitAndLossStatementDrillThrough":
                case "BalanceSheetDrillthrough":
                case "BalanceSheetComp":
                case "ProfitAndLossComp":
                case "BalanceSheetPropBreakdown":
                case "ProfitAndLossExtendedVariance":
                    return new RowSegmentFormulaGenerator();



                case "ProfitAndLossBudget":
                    return new ProfitAndLossBudgetFormulas();



                case "TrialBalanceVariance":
                    return new TrialBalanceVarianceGenerator();



                case "ProfitAndLossStatementByJob":
                    RowSegmentFormulaGenerator gen = new RowSegmentFormulaGenerator();
                    gen.trimFormulaRange = false;
                    return gen;




                case "PayablesAccountReport":
                    return new MultiFormulaGenerator(new RowSegmentFormulaGenerator(), new SumWithinSegmentGenerator());



                case "ReportOutstandingBalance":
                    switch (worksheetNum)
                    {
                        case 0:
                            return new MultiFormulaGenerator(new PeriodicFormulaGenerator(), new SumOtherSums());
                        default:
                            return new FullTableFormulaGenerator();
                    }



                case "RentRollActivityItemized":
                case "RentRollActivityItemized_New":
                    PeriodicFormulaGenerator mainFormulas = new PeriodicFormulaGenerator();
                    mainFormulas.SetDataCellDefenition(cell => FormulaManager.IsEmptyCell(cell) || FormulaManager.IsDollarValue(cell));

                    SumOtherSums otherFormulas = new SumOtherSums();

                    return new MultiFormulaGenerator(mainFormulas, otherFormulas);



                case "RentRollHistory":
                    switch (worksheetNum)
                    {
                        case 0:
                            return new RentRollHistorySheet1();
                        case 1:
                            return new RentRollHistorySheet2();
                        default:
                            return null;
                    }



                case "VendorInvoiceReportWithJournalAccounts":
                    IFormulaGenerator generatorToBeUsed = null;

                    var worksheetType = MetaDataGathering.DetectWorksheetType(workbook.Worksheets[worksheetNum]);


                    switch (worksheetType)
                    {
                        case MetaDataGathering.VendorInvoiceReportWithJAType.BLANK:
                            return null;

                        case MetaDataGathering.VendorInvoiceReportWithJAType.STANDARD:
                            generatorToBeUsed = new VendorInvoiceReportFormulas();
                            break;

                        case MetaDataGathering.VendorInvoiceReportWithJAType.CHART:
                            generatorToBeUsed = new FullTableFormulaGenerator();
                            break;


                        default:
                            Console.WriteLine($"Error: worksheet structure unrecongnized. Unable to add formulas to worksheet {worksheetNum + 1}");
                            Console.WriteLine("You may be able to fix this by finding the appropriate formula generator for this worksheet.");
                            return null;
                    }


                    //we need to set the BeyondFormulRange of SumOtherSums to allow "invoice Total:"
                    SumOtherSums sumOfSums = new SumOtherSums();
                    sumOfSums.SetDefenitionForBeyondFormulaRange(cell => !FormulaManager.IsEmptyCell(cell) 
                                                                        && !FormulaManager.IsDollarValue(cell) 
                                                                        && cell.Text.Trim() != "Invoice Total:");

                    return new MultiFormulaGenerator(new PeriodicFormulasOnTop(), generatorToBeUsed, sumOfSums);





                //This report needs to be done a little differently becuase it doesnt always have the same number of worksheets
                case "RentRollAllItemized":
                    if (worksheetNum == workbook.Worksheets.Count - 1) //for the last worksheet
                    {
                        //This report has multiple tables that are organized in different ways. We need 
                        //two different formula generators to ensure that all tables get done correctly

                        FullTableFormulaGenerator first = new FullTableFormulaGenerator();
                        RowSegmentFormulaGenerator second = new RowSegmentFormulaGenerator();
                        IsDataCell dataCellDef = new IsDataCell(cell =>
                                FormulaManager.IsDollarValue(cell)
                                || FormulaManager.IsIntegerWithCommas(cell)
                                || FormulaManager.IsPercentage(cell)
                                || FormulaManager.CellHasFormula(cell));


                        first.SetDefenitionForBeyondFormulaRange(first.IsNonDataCell);

                        MultiFormulaGenerator generator = new MultiFormulaGenerator(first, second);
                        generator.SetDataCellDefenition(dataCellDef);
                        return generator;
                    } 
                    else if (worksheetNum == 1)
                    {
                        return new MultiFormulaGenerator(new PeriodicFormulaGenerator(), 
                            new SumOtherSums(), new FormulaBetweenSheets());
                    }

                    else
                    {
                        return new MultiFormulaGenerator(new PeriodicFormulaGenerator(), new SumOtherSums());
                    }
                            
                    




                case "RentRollAll":
                    switch (worksheetNum)
                    {
                        case 1:
                            //This report has multiple tables that are organized in different ways. We need 
                            //two different formula generators to ensure that all tables get done correctly

                            FullTableFormulaGenerator first = new FullTableFormulaGenerator();
                            RowSegmentFormulaGenerator second = new RowSegmentFormulaGenerator();
                            IsDataCell dataCellDef = new IsDataCell(cell =>
                                    FormulaManager.IsDollarValue(cell)
                                    || FormulaManager.IsIntegerWithCommas(cell)
                                    || FormulaManager.IsPercentage(cell)
                                    || FormulaManager.CellHasFormula(cell));


                            first.SetDefenitionForBeyondFormulaRange(first.IsNonDataCell);

                            MultiFormulaGenerator generator = new MultiFormulaGenerator(first, second);
                            generator.SetDataCellDefenition(dataCellDef);
                            return generator;



                        default:
                            var fullTableGen = new FullTableFormulaGenerator();
                            fullTableGen.SetDefenitionForBeyondFormulaRange(fullTableGen.IsNonDataCell);
                            return fullTableGen;


                    }



                case "VacancyLoss":
                    switch (worksheetNum)
                    {
                        case 0:
                            formulaGenerator = new FullTableFormulaGenerator();
                            int ignoredOutput;
                            formulaGenerator.SetDataCellDefenition(cell => FormulaManager.IsDollarValue(cell) || Int32.TryParse(cell.Text, out ignoredOutput));
                            return formulaGenerator;

                        default:
                            return new FullTableFormulaGenerator();
                    }



                case "ReportCashReceipts":
                    return new ReportCashRecipts();



                case "ChargesCreditsReport":
                    return new ChargesCreditReportFormulas();




                case "RentRollActivityCompSummary":
                //case "SubsidyRentRollReport": //it has been decided that this report doeant get any formulas
                    return new SummaryColumnGenerator();




                case "ReportPayablesRegister":
                case "AgedPayables":
                    formulaGenerator = new FullTableFormulaGenerator();
                    formulaGenerator.SetDefenitionForBeyondFormulaRange(formulaGenerator.IsNonDataCell);
                    return formulaGenerator;




                case "AgedReceivables":
                    //This report sometimes has numerous rows with subtotals and sometimes does not
                    if (MetaDataGathering.AgedReceivablesNeedsSubtotals(workbook.Worksheets[0]))
                    {
                        formulaGenerator = new SumOtherSums();
                        formulaGenerator.SetDefenitionForBeyondFormulaRange(cell => !FormulaManager.IsDollarValue(cell) 
                                                                                 && !FormulaManager.IsEmptyCell(cell)
                                                                                 && !FormulaManager.CellHasFormula(cell));


                        PeriodicFormulaGenerator periodic = new PeriodicFormulaGenerator();
                        periodic.SetSummaryCellDefenition(cell => cell.Style.Font.Bold && 
                                        (FormulaManager.IsDollarValue(cell) || FormulaManager.CellHasFormula(cell)));

                        return new MultiFormulaGenerator(formulaGenerator, periodic);
                    }
                    else
                    {
                        formulaGenerator = new FullTableFormulaGenerator();
                        formulaGenerator.SetDefenitionForBeyondFormulaRange(formulaGenerator.IsNonDataCell);
                        return formulaGenerator;
                    }
                    




                case "TrialBalance":
                    formulaGenerator = new SumOnlyBolds();
                    formulaGenerator.SetDefenitionForBeyondFormulaRange(formulaGenerator.IsNonDataCell);
                    return formulaGenerator;




                case "CollectionsAnalysisSummary":
                    formulaGenerator = new FullTableFormulaGenerator();
                    formulaGenerator.SetDataCellDefenition(                                     //matches a percentage
                        cell => FormulaManager.IsDollarValue(cell) || Regex.IsMatch(cell.Text, "^\\d?\\d{2}([.]\\d{2})?%$"));


                    return formulaGenerator;



                case "RentRollPortfolio":
                    formulaGenerator = new FullTableFormulaGenerator();
                    double ignored;
                    formulaGenerator.SetDataCellDefenition(cell => FormulaManager.IsDollarValue(cell) || Double.TryParse(cell.Text, out ignored));
                    formulaGenerator.SetDefenitionForBeyondFormulaRange(formulaGenerator.IsNonDataCell);
                    return formulaGenerator;



                case "ProfitAndLossStatementByPeriod":
                    FullTableSummaryColumn summaryCol = new FullTableSummaryColumn();
                    summaryCol.SetStopGivingFormulas(cell => !FormulaManager.IsEmptyCell(cell) 
                                                            && !FormulaManager.IsDollarValue(cell));


                    return new MultiFormulaGenerator(summaryCol, new FullTableFormulaGenerator());




                case "UnitInvoiceReport":
                case "VendorInvoiceReport":
                    return new MultiFormulaGenerator(new PeriodicFormulasOnTop(), new SumOtherSums());



                case "ReportCashReceiptsSummary":
                    //seperate formula generator to add the monthly sections of the report
                    SumWithinSegmentGenerator sumEachMonth = new SumWithinSegmentGenerator();
                    sumEachMonth.UseArrayFormula = false;
                    sumEachMonth.SumNonFormulas = false;

                    return new MultiFormulaGenerator(new FullTableFormulaGenerator(), sumEachMonth);




                case "ReportAccountBalances":
                case "ReportTenantBal":
                case "LedgerReport":
                case "RentRollActivity_New":
                case "RentRollActivity":
                case "JournalLedger":
                case "AgedAccountsReceivable":
                case "CollectionsAnalysis":
                case "InvoiceRecurringReport":
                    return new FullTableFormulaGenerator();





                case "RentRollActivityTotals":
                    FullTableFormulaGenerator g = new FullTableFormulaGenerator();
                    g.SetDataCellDefenition(
                        cell => FormulaManager.IsDollarValue(cell) || FormulaManager.IsIntegerValue(cell));

                    return g;




                //These reports dont fit into any existing system
                //AgedAccountsReceivable (its original totals are incorrect)




                //Reports Im working on
                case "Budget":
                    return new MultiFormulaGenerator(new FullTableSummaryColumn(), new RowSegmentFormulaGenerator());
                case "RentRollCommercialItemized":


                //problem: not sure what to add up here
                case "ReportEscalateCharges":




                //This report cannot get formulas because it does not include some necessary data
                case "PaymentsHistory":



                default:
                    return null;
            }
        }



        /// <summary>
        /// Retrieves the required arguments that should be passed into IFormulaGenerator.InsertFormulas function
        /// for a given report and worksheet.
        /// </summary>
        /// <param name="reportName">the name of the report getting the formulas</param>
        /// <param name="worksheetNum">the index of the worksheet getting the formulas</param>
        /// <param name="workbook">The workbook being given formulas. Sometimes needed to tell how many worksheets there are</param>
        /// <param name="reportVersion">teh version of the report being unmerged. Null or empty if only one version exists</param>
        /// <returns>
        /// a list of strings that should be passed to the formula generator when formulas are being added,
        /// or null if the worksheet does not require formulas
        /// </returns>
        internal static string[] GetFormulaGenerationArguments(string reportName, int worksheetNum, ExcelWorkbook workbook, string reportVersion)
        {
            switch (reportName)
            {


                case "ProfitAndLossStatementByPeriod":
                    return new string[] { "1Total", "2Total Income", "2Total Expense", "2Total Non-Operating Income",
                        "2Total Other Cash Adjustments", "Net Operating Income~-Total Expense,Total Income",
                        "Net Income~+Total Income,+-Total Expense",
                        "Adjusted Net Income~Total Non-Operating Income,Total Other Cash Adjustments,Net Operating Income" };




                case "BalanceSheetDrillthrough":
                    return new string[] { "Asset=Total Asset", "Current Assets=Total Current Assets",
                        "Fixed Asset=Total Fixed Asset", "Other Asset=Total Other Asset", "Current Liabilities=Total Current Liabilities",
                        "Liability=Total Liability", "Long Term Liability=Total Long Term Liability", 
                        "Equity=Total Equity", "Total Liabilities~Total Long Term Liability,Total Liability,Total Current Liabilities",
                        "Total Assets~Total Other Asset,Total Fixed Asset,Total Current Assets", 
                        "Total Liabilities And Equity~Total Equity,Total Liabilities" };




                case "ProfitAndLossComp":
                    return new string[] { "INCOME=Total Income", "EXPENSE=Total Expense",
                        "Non-Operating Income=Total Non-Operating Income",
                        "Other Cash Adjustments=Total Other Cash Adjustments",
                        "Net Operating Income~Total Income,-Total Expense",
                        "Net Income~+Total Income,-+Total Expense",
                        "Adjusted Net Income~Total Other Cash Adjustments,Total Non-Operating Income,Net Operating Income" };



                case "RentRollActivity":
                case "RentRollActivity_New":
                    switch (worksheetNum)
                    {
                        case 0:
                            return new string[] { "Total:" };

                        case 1:
                            return new string[] { $"Total For {anyProperty}:" };

                        default:
                            return new string[0];
                    }




                case "ReportCashReceiptsSummary":
                    switch (worksheetNum)
                    {
                        case 0:
                            return new string[] { "1Total Tenant Receivables:", "1Total Other Receivables:", 
                                $"2{anyMonth} {anyYear}=Total For {anyMonth} {anyYear}:",
                                $"Total For {anyProperty}:~+Total For {anyMonth} [12]\\d\\d\\d:" };

                        default:
                            return new string[] { "1Total Tenant Receivables:" };
                    }




                case "ReportTenantBal":
                    return new string[] { "Total Open Charges:", 
                        "Balance:~Total Open Charges:,Total Future Charges:,Total Unallocated Payments:" };




                case "ProfitAndLossBudget":
                    return new string[] { "INCOME=Total Income", "EXPENSE=Total Expense", 
                        "Net Operating Income~Total Income,-Total Expense", 
                        "Net Income~-Total Expense,Total Income,Net Operating Income" };




                case "BalanceSheetComp":
                    return new string[] { "Current Assets=Total Current Assets", "Fixed Asset=Total Fixed Asset",
                        "Other Asset=Total Other Asset", "Current Liabilities=Total Current Liabilities",
                        "Liability=Total Liability", "Liabilities And Equity=Total Liabilities And Equity",
                        "Long Term Liability=Total Long Term Liability", "Equity=Total Equity", 
                        "Total Liabilities~Total Long Term Liability,Total Liability,Total Current Liabilities",
                        "Total Assets~Total Other Asset,Total Fixed Asset,Total Current Assets" };




                case "ChargesCreditsReport":
                    return new string[] { "Total: \\$(\\d\\d\\d,)*\\d?\\d?\\d[.]\\d\\d" };



                /*
                 //* It has been decided that this report doesnt get any formulas
                case "SubsidyRentRollReport":
                    return new string[] { 
                        "Current Tenant \\sPortion of the Rent,Current  Subsidy Portion of the Rent=>Current Monthly \\sContract Rent" };
                */



                case "RentRollActivityCompSummary":
                    return new string[] { "-Opening A/R,Closing A/R=>A/R [+][(]-[)]" };




                case "RentRollHistory":
                    switch (worksheetNum)
                    {
                        case 1:
                            return new string[] { "Residential: \\$\\d+(,\\d\\d\\d)*[.]\\d\\d", "Total: \\$\\d+(,\\d\\d\\d)*[.]\\d\\d", };

                        default:
                            return new string[0];
                    }






                case "RentRollActivityItemized":
                case "RentRollActivityItemized_New":
                    //Old regex: r=(\\d{4})|([A-Z]\\d\\d)
                    return new string[] { "1r=[A-Z0-9]+", "1Beg\\s+Balance", "1Charges", "1Adjustments",
                        "1Payments", "1End Balance", "1Change", "2Total:" };




                case "BalanceSheetPropBreakdown":
                    return new string[] { "Current Assets=Total Current Assets", "Fixed Asset=Total Fixed Asset",
                        "Other Asset=Total Other Asset", "Current Liabilities=Total Current Liabilities", 
                        "Long Term Liability=Total Long Term Liability", "Equity=Total Equity", 
                        "Total Assets~Total Other Asset,Total Fixed Asset,Total Current Assets", 
                        "Total Liabilities~Total Current Liabilities,Total Long Term Liability", 
                        "Total Liabilities And Equity~Total Equity,Total Liabilities" };




                case "VendorInvoiceReportWithJournalAccounts":
                    var worksheetType = MetaDataGathering.DetectWorksheetType(workbook.Worksheets[worksheetNum]);

                    switch (worksheetType)
                    {
                        case MetaDataGathering.VendorInvoiceReportWithJAType.STANDARD:
                            return new string[] { "1Amount Owed", "1Amount Paid", "1Balance",
                                                "2Amount Owed", "2Amount Paid", "2Balance", "3Total:" };

                        case MetaDataGathering.VendorInvoiceReportWithJAType.CHART:
                            return new string[] { "1Amount Owed", "1Amount Paid", "1Balance", "2Total:", "3Total:" };

                        default:
                            return new string[0];
                    }




                case "ReportCashReceipts":
                    return new string[] { "r=[A-Z]\\d{4}", "Charge Total", "Amount" };




                //This report needs to be done a little differently becuase it differs in the number of worksheets
                case "RentRollAllItemized":
                    if(worksheetNum == workbook.Worksheets.Count - 1) //if we are on the last worksheet
                    {
                        return new string[] { "1Total:", "2Subtotals=Total:" };
                    }
                    else if(worksheetNum == 1)
                    {
                        return new string[] { "1r=[A-Z]-\\d\\d", "1Monthly Charge", "1Annual Charge", "2Total:", "3sheet0", "3sheet1" };
                    }
                    else if(worksheetNum == 0)
                    {
                        return new string[] { "1r=[A-Z]-\\d\\d", "1Monthly Charge", "1Annual Charge", "2Total:" };
                    }
                    else
                    {
                        return new string[0];
                    }





                case "RentRollAll":
                    switch (worksheetNum)
                    {
                        case 1:
                            return new string[] { "1Total:", "2Subtotals=Total:" };

                        default:
                            return new string[] { "Total:" };
                    }



                case "ProfitAndLossStatementDrillthrough":
                case "ProfitAndLossStatementDrillThrough":
                    return new string[] { "Expense=Total Expense", "Income=Total Income",
                        "Non-Operating Income=Total Non-Operating Income",
                        "Other Cash Adjustments=Total Other Cash Adjustments",
                        "Net Operating Income~-Total Expense,Total Income",
                        "Net Income~+Total Income,+-Total Expense",
                        "Adjusted Net Income~Total Other Cash Adjustments,Total Non-Operating Income,Net Operating Income" };




                case "PayablesAccountReport":
                    return MetaDataGathering.GetHeadersForPayablesAccountReport(workbook.Worksheets[worksheetNum]);


                    //old version
                    /*
                    return new string[] { "1Pool Furniture=Total Pool Furniture", "1Hallways=Total Hallways", 
                        "1Garage=Total Garage", "1Elevators=Total Elevators", "1Clubhouse=Total Clubhouse",
                        "1Painting=Total Painting", "1HVAC=Total HVAC", "1Windows=Total Windows", "1Appliances=Total Appliances",
                        "1Paint/Contracting Labor=Total Paint/Contracting Labor",
                        "2Common Area CapEx=Total Common Area CapEx", "2CapEx=Total CapEx",
                        "2Apartment Renovation=Total Apartment Renovation",
                        "Total~Total Common Area CapEx,Total CapEx,Total Apartment Renovation",
                        "Total:~Total Common Area CapEx,+Total CapEx,Total Apartment Renovation" };
                        */




                case "ReportOutstandingBalance":
                    switch (worksheetNum)
                    {
                        case 0:
                            return new string[] { "1r=[A-Z0-9]+", "1Balance", "2Total For ([A-Z][a-z]+)( [A-Z]?[a-z]+)+:" };

                        default:
                            return new string[] { "Total" };
                    }




                case "CollectionsAnalysis":
                case "ReportPayablesRegister":
                case "AgedAccountsReceivable":
                case "ReportAccountBalances":
                case "JournalLedger":
                case "CollectionsAnalysisSummary":
                case "AgedPayables":
                    return new string[] { "Total" };




                case "AgedReceivables":
                    if (MetaDataGathering.AgedReceivablesNeedsSubtotals(workbook.Worksheets[0]))
                    {
                        var headers = MetaDataGathering
                            .GetColumnNamesForAgedReceivablesSubtotals(workbook.Worksheets[0]);
                            


                        List<string> args = new List<string>();
                        args.Add("1Total");
                        args.Add("2r=[A-Z]?\\d{1,4}[A-Z]?");
                        args.AddRange(headers);


                        return args.ToArray();


                        //old solution
                        /*return new string[] { "1Total", "2r=\\d{1,4}[A-Z]\\-[A-Z]{1,4}", "20 - 30", "231 - 60",
                                                "261 - 90", "2Over 90", "2Total"};*/
                                             
            }
                    else
                    {
                        return new string[] { "Total" };
                    }




                case "VacancyLoss":
                case "InvoiceRecurringReport":
                case "RentRollPortfolio":
                case "TrialBalance":
                    return new string[] { "Total:" };



                case "UnitInvoiceReport":
                case "VendorInvoiceReport":
                    return new string[] { "1Amount Owed", "1Amount Paid", "1Balance", "2Total:" };



                case "ProfitAndLossStatementByJob":
                    return new string[] { "Income=Total Income", "Expense=Total Expense",
                        "Other Cash Adjustments=Total Other Cash Adjustments", "Non-Operating Income=Total Non-Operating Income", 
                        "Net Income~Total Income,-Total Expense", "Net Operating Income~Total Income,-Total Expense",
                        "Adjusted Net Income~Total Other Cash Adjustments,Total Non-Operating Income,Net Operating Income" };



                case "TrialBalanceVariance":
                    return new string[] { "Asset=Total Asset", "Current Assets=Total Current Assets",
                        "Liability=Total Liability", "Current Liabilities=Total Current Liabilities",
                        "Equity=Total Equity", "Income=Total Income", "Expense=Total Expense",
                        "Total:~Total Expense,Total Income,Total Equity,Total Liability,Total Asset" };



                case "ProfitAndLossExtendedVariance":
                    return new string[] { "INCOME=Total Income", "EXPENSE=Total Expense",
                        "Net Operating Income~Total Income,-Total Expense", 
                        "Net Income~Net Operating Income,Total Income,-Total Expense" };



                case "LedgerReport":
                    return MetaDataGathering.GetHeadersForLedgerReport(workbook.Worksheets[worksheetNum]);


                case "RentRollActivityTotals":
                    return new string[] { "Totals For All Buildings" };



                //Reports with minor issues:
                //ProfitAndLossExtendedVariance
                //AgedAccountsReceivable
                //RentRollCommercialItemized
                //LedgerReport




                //these reports I'm still working on
                case "Budget":
                    return new string[] { "1Total", "2INCOME=TOTAL INCOME", "2EXPENSE=TOTAL EXPENSE" };
                    //FIXME: I am not sure what rows need formulas
                case "RentRollCommercialItemized": //not sure what Im supposed to be adding here
                case "ReportEscalateCharges": //problem: not sure what to add up




                // this report does not have the necessary columns/data to get a formula
                // for the time being this report gets no formulas
                case "PaymentsHistory": 



                default:
                    return new string[0];
            }
        }
    }
}
