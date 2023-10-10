

namespace ExcelCleanerNet45
{


    /// <summary>
    /// Enum for tracking what kind of merge a cell has
    /// </summary>
    public enum MergeType
    {
        NOT_A_MERGE,                //no merge in the address
        EMPTY,                      //merge cell with no text
        MAIN_HEADER,                //merge cell outside the table with a header (ussually) describing the table contents
        MINOR_HEADER,               //merge cell inside the table with a header (usually) describing row contents
        DATA                        //merge cell containing data
    }
}

