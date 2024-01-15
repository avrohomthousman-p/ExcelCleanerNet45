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
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-5\LegalReport_7232023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-5\ReportTenantBal_7232023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-5\RentRollActivity_New_7232023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-5\RentRollActivityCompSummary_7232023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-5\RentRollHistory_7232023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-5\ReportOutstandingBalance_7232023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-5\ReportCashReceiptsSummary_7232023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-5\ReportCashReceipts_7232023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-5\ProfitAndLossStatementByPeriod_7232023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-5\PayablesAccountReport_7232023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-5\ReportPayablesRegister_7242023.xlsx
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
                // C:\Users\avroh\Downloads\ExcelProject\missing-reports\DistributionsReport.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\missing-reports\CollectionsAnalysis.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\missing-reports\InvoiceRecurringReport.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\missing-reports\ProfitAndLossExtendedVariance.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\missing-reports\VendorInvoiceReport.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\missing-reports\ReportPayablesRegister.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\missing-reports\ProfitAndLossStatementByJob.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\missing-reports\UnitInvoiceReport.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\missing-reports\TrialBalanceVariance.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\missing-reports\RentRollActivityTotals.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\missing-reports\TrialBalance.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\bugged-reports\RentRollActivityCompSummary.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\bugged-reports\RentRollHistory.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\bugged-reports\AgedReceivables.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\bugged-reports\RentRollAll.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\bugged-reports\RentRollAllItemized.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\bugged-reports\RentRollActivity_New.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\bugged-reports\RentRollBalanceHistory.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\bugged-reports\ReportOutstandingBalance.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\bugged-reports\ReportCashReceiptsSummary.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\bugged-reports\ChargesCreditsReport.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\bugged-reports\ProfitAndLossStatementDrillthrough_3.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\bugged-reports\ProfitAndLossStatementByPeriod_2.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\bugged-reports\ProfitAndLossBudget.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\bugged-reports\ProfitAndLossComp_3.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\bugged-reports\ProfitAndLossStatementByJob.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\bugged-reports\TrialBalance_2.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\bugged-reports\SummaryReport.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\bugged-reports\ReportTenantSummary.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\bugged-reports\BalanceSheetDrillthrough.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\bugged-reports\CCTransactionsReport.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\bugged-reports\InvoiceDetail.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\bugged-reports\UnitInvoiceReport.xlsx
                // C:\Users\avroh\Downloads\UnitInvoiceReport JustAmounts_1.xlsx





                Console.WriteLine("Please enter the filepath (or directory) of the Excel report you want to clean:");
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



            //Are we doing a single file or an entire directory
            if (!filepath.Contains("."))
            {
                RunAllReportsInDirectory(filepath);
            }
            else
            {
                Tuple<string, string> reportData = GetReportNameAndVersion(filepath);

                //Tell the file cleaner to do the cleaning
                byte[] output = FileCleaner.OpenXLSX(ConvertFileToBytes(filepath), reportData.Item1, reportData.Item2, true);


                //save the output
                SaveByteArrayAsFile(output, filepath.Replace(".xlsx", "_fixed.xlsx"));
            }


            
            Console.WriteLine("Press Enter to exit");
            Console.Read();

        }



        /// <summary>
        /// Extracts the report name and version from the file path given
        /// </summary>
        /// <param name="filename">the name of the report's file</param>
        /// <returns>the name of the report type and the report version</returns>
        private static Tuple<string,string> GetReportNameAndVersion(string filename)
        {

            int start = filename.LastIndexOf('\\') + 1;
            int length;


            //First remove the numbers and .xlsx at the end of the file name (and the full file path)

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




            //now seperate the name from the version if the version is present

            filename = filename.Substring(start, length);
            int whitespace = filename.IndexOf(' ');

            if(whitespace < 0)
            {
                return new Tuple<string, string>(filename, "");
            }
            else
            {
                return new Tuple<string, string>(filename.Substring(0, whitespace), filename.Substring(whitespace + 1));
            }
        }




        /// <summary>
        /// Cleans all reports in the specified directory
        /// </summary>
        /// <param name="directory">the full path of the directory containing the reports</param>
        private static void RunAllReportsInDirectory(string directory)
        {
            DirectoryInfo d = new DirectoryInfo(directory);
            foreach (FileInfo file in d.EnumerateFiles())
            {
                //ensure that it is an excel file and not some other file
                if(!file.Name.EndsWith(".xlsx") && !file.Name.EndsWith(".xls"))
                {
                    continue;
                }

                //ensure that it hasnt already been cleaned in a previous run
                Regex regex = new Regex("^.+_[Ff]ixed[.]xls(x)?$"); //matches if the report name has a "fixed"                                                     
                if (regex.IsMatch(file.Name))
                {
                    continue;
                }



                Tuple<string, string> reportData = GetReportNameAndVersion(file.FullName);                


                Console.WriteLine("cleaning report " + file.Name);

                //Tell the file cleaner to do the cleaning
                byte[] output = FileCleaner.OpenXLSX(ConvertFileToBytes(file.FullName), reportData.Item1, reportData.Item2, true);


                //save the output
                SaveByteArrayAsFile(output, file.FullName.Replace(".xlsx", "_fixed.xlsx"));

            }
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
