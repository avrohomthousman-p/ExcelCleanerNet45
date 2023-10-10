using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ExcelCleanerNet45
{
    class Program
    {

        /// <summary>
        /// The main entry point for the application when you want to clean a file saved on your computer.
        /// For a web based cleaning, just call the FileCleaner.OpenXlsx method directly.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            string filepath = "";

            if (args != null && args.Count() > 0)
            {
                filepath = args[0];
            }

            else
            {

                // C:\Users\avroh\Downloads\ExcelProject\PayablesAccountReport_large.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\PayablesAccountReport_1Prop.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\ReportPayablesRegister.xlsx

                // C:\Users\avroh\Downloads\ExcelProject\ProfitAndLossStatementDrillthrough.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\AgedReceivables.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\LedgerExport.xlsx

                // C:\Users\avroh\Downloads\ExcelProject\TrialBalance.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\ProfitAndLossStatementByPeriod.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\testFile.xlsx

                // C:\Users\avroh\Downloads\ExcelProject\BalanceSheetComp_742023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\BalanceSheetDrillthrough_722023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\BankReconcilliation_722023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\PaymentsHistory_722023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\LedgerReport_722023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports\AgedReceivables_7102023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports\BalanceSheetComp_7102023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports\AdjustmentReportMult_7102023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-2\AdjustmentReport_7102023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-3\CashFlow_7182023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-3\ChargesCreditsReport_7182023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-3\ProfitAndLossBudget_7182023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-3\CreditCardStatement_7182023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-3\CollectionsAnalysisSummary_7182023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-4\AgedAccountsReceivable_7192023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-4\BankReconcilliation_7192023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-4\ChargesCreditsReport_7192023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-4\BalanceSheetPropBreakdown_7192023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-4\BalanceSheetComp_7192023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-4\AgedReceivables_7192023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-4\AgedPayables_7192023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-5\BalanceSheetComp_7232023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-5\AgedAccountsReceivable_7232023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-5\BalanceSheetDrillthrough_7232023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-5\AdjustmentReport_7232023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-5\AgedReceivables_7232023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-5\ProfitAndLossStatementDrillThrough_7232023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-5\RentRollAll_7232023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-5\InvoiceDetail_7232023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-5\LedgerReport_7232023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-5\ReportTenantBal_7232023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-5\RentRollActivity_New_7232023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-5\RentRollActivityCompSummary_7232023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-5\RentRollHistory_7232023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-5\ReportOutstandingBalance_7232023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-5\ReportCashReceiptsSummary_7232023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-5\ReportCashReceipts_7232023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-5\ProfitAndLossStatementByPeriod_7232023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-5\PayablesAccountReport_7232023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-5\PendingWebPayments_7232023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-5\ProfitAndLossBudget_7232023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-5\ProfitAndLossComp_7232023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-5\ChargesCreditsReport_7232023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-5\RentRollPortfolio_7232023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-5\PreprintedLeasesReport_7232023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-6\ReportTenantSummary_7252023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-6\TenantDirectory_7252023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-6\VacancyLoss_7252023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-6\SubsidyRentRollReport_7252023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-6\VendorInvoiceReportWithJournalAccounts_7252023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\missing-reports\JournalLedger_8222023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\missing-reports\RentRollActivity_New_8222023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\missing-reports\RentRollActivityItemized_New_8222023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\missing-reports\ReportAccountBalances_8222023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\missing-reports\RentRollAll.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\missing-reports\RentRollAllItemized_8222023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\missing-reports\Budget.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\missing-reports\Budget_2.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\missing-reports\CollectionsAnalysis.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\missing-reports\InvoiceRecurringReport.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\missing-reports\ProfitAndLossExtendedVariance.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\missing-reports\VendorInvoiceReport.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\missing-reports\ReportPayablesRegister.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\missing-reports\ProfitAndLossStatementByJob.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\missing-reports\UnitInvoiceReport.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\missing-reports\TrialBalanceVariance.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\missing-reports\RentRollActivityTotals.xlsx




                Console.WriteLine("Please enter the filepath of the Excel report you want to clean:");
                filepath = Console.ReadLine();

                /*
                OpenFileDialog dialog = new OpenFileDialog();
                if (DialogResult.OK == dialog.ShowDialog())
                {
                    filepath = dialog.FileName;
                }
                Console.WriteLine("Hello World!");
                */
            }


            string reportName = GetReportName(filepath);

            //Tell the file cleaner to do the cleaning
            byte[] output = FileCleaner.OpenXLSX(ConvertFileToBytes(filepath), reportName, true);


            //save the output
            SaveByteArrayAsFile(output, filepath.Replace(".xlsx", "_fixed.xlsx"));
            Console.WriteLine("Press Enter to exit");
            Console.Read();

        }



        /// <summary>
        /// Extracts the report name from the file path given
        /// </summary>
        /// <param name="filename">the name of the report's file</param>
        /// <returns>the name of the report type</returns>
        private static string GetReportName(string filename)
        {

            int start = filename.LastIndexOf('\\') + 1;
            int length;



            Regex regex = new Regex("^.+(_\\d+)[.]xlsx$"); //matches if the report name ends with an underscore followed by numbers

            if (regex.IsMatch(filename))
            {
                length = filename.Length - start;
                length -= (filename.Length - filename.LastIndexOf('_')); //minus the number of characters after the file name
            }
            else
            {
                length = filename.Length - start - 5; //if we just need to remove the .xlsx at the end
            }




            return filename.Substring(start, length);
        }




        /// <summary>
        /// Opens the specified file and writes its contents to a byte array. This function is only needed for testing. In production
        /// the file itself will be passed in as a byte array, not as a filepath.
        /// </summary>
        /// <param name="filepath">the location of the file</param>
        /// <returns>a byte array with the contents of the file in it</returns>
        private static byte[] ConvertFileToBytes(string filepath)
        {
            FileInfo existingFile = new FileInfo(filepath);
            byte[] fileData = new byte[existingFile.Length];


            var fileStream = existingFile.Open(FileMode.Open);
            int bytesRead = 0;
            int bytesToRead = (int)existingFile.Length;
            while (bytesToRead > 0)
            {
                int justRead = fileStream.Read(fileData, bytesRead, bytesToRead);

                if (justRead == 0)
                {
                    break;
                }

                bytesRead += justRead;
                bytesToRead -= justRead;
            }


            fileStream.Close();
            return fileData;
        }




        /// <summary>
        /// Saves the specified byte array to a file.
        /// </summary>
        /// <param name="fileData">the byte array that should be saved to the file</param>
        /// <param name="filepath">the filepath of the file</param>
        private static void SaveByteArrayAsFile(byte[] fileData, string filepath)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;


            using (ExcelPackage package = new ExcelPackage(new MemoryStream(fileData)))
            {
                package.SaveAs(filepath);
            }
        }
    }
}
