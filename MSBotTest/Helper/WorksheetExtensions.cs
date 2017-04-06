using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;

namespace MSBotTest
{
    /// <summary>
    /// worksheet extension class
    /// </summary>
    public static class WorksheetExtensions
    {
        /// <summary>
        /// Get the cell value of address for the sheet
        /// code taken from https://msdn.microsoft.com/en-us/library/office/hh298534.aspx and modified
        /// </summary>
        /// <param name="sheetName"></param>
        /// <param name="addressName"></param>
        /// <returns></returns>
        public static string GetCellValue(this Sheet sheet, string addressName)
        {
            string value = null;
            //Cell cell = sheet.Elements<Cell>().Where(c => c.CellReference == addressName).FirstOrDefault();
            // get the workbookpart
            WorkbookPart wbPart = sheet.Ancestors<Workbook>().First().WorkbookPart;

            // Retrieve a reference to the worksheet part.
            WorksheetPart wsPart =
                (WorksheetPart)(wbPart.GetPartById(sheet.Id));

            // Use its Worksheet property to get a reference to the cell 
            // whose address matches the address you supplied.
            Cell theCell = wsPart.Worksheet.Descendants<Cell>().
              Where(c => c.CellReference == addressName).FirstOrDefault();

            // If the cell does not exist, return an empty string.
            if (theCell != null)
            {
                if (theCell.CellValue == null) return string.Empty;

                value = theCell.CellValue.InnerText;

                // If the cell represents an integer number, you are done. 
                // For dates, this code returns the serialized value that 
                // represents the date. The code handles strings and 
                // Booleans individually. For shared strings, the code 
                // looks up the corresponding value in the shared string 
                // table. For Booleans, the code converts the value into 
                // the words TRUE or FALSE.
                if (theCell.DataType != null)
                {
                    switch (theCell.DataType.Value)
                    {
                        case CellValues.SharedString:

                            // For shared strings, look up the value in the
                            // shared strings table.
                            var stringTable =
                                wbPart.GetPartsOfType<SharedStringTablePart>()
                                .FirstOrDefault();

                            // If the shared string table is missing, something 
                            // is wrong. Return the index that is in
                            // the cell. Otherwise, look up the correct text in 
                            // the table.
                            if (stringTable != null)
                            {
                                value =
                                    stringTable.SharedStringTable
                                    .ElementAt(int.Parse(value)).InnerText;
                            }
                            break;

                        case CellValues.Boolean:
                            switch (value)
                            {
                                case "0":
                                    value = "FALSE";
                                    break;
                                default:
                                    value = "TRUE";
                                    break;
                            }
                            break;
                    }
                }
            }
            return value;
        }
    }
}
