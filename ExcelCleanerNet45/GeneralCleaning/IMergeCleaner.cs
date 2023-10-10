using OfficeOpenXml;

namespace ExcelCleanerNet45
{
    internal interface IMergeCleaner
    {

        /// <summary>
        /// Unmerges all merges in the specified worksheet and ensures little to no formatting is lost in the process
        /// </summary>
        /// <param name="worksheet">the worksheet that needs its merged cells unmerged</param>
        /// <exception cref="">If something goes wrong when trying to unmerge the merged cells</exception>
        void Unmerge(ExcelWorksheet worksheet);
    }
}

