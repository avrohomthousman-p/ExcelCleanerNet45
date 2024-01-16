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
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-5\BalanceSheetComp_7232023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-6\ReportTenantSummary_7252023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\missing-reports\InvoiceList.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\missing-reports\TrialBalanceVariance.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\bugged-reports\InvoiceList.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\bugged-reports\UnitInvoiceReport.xlsx
                // C:\Users\avroh\Downloads\ProfitAndLossStatementByJob.xlsx





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
