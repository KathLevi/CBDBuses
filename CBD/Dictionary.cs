using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace CBD
{
    class NameAndLocation
    {

        public static Excel.Application my_excel = new Excel.Application();
        public static Excel.Workbook my_book = my_excel.Workbooks.Open(@"CBD\Locations.xlsx", 0, false, 5, "", "",
            false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
        public static Excel.Sheets my_sheets = my_book.Worksheets;
        public static Excel.Worksheet locations = (Excel.Worksheet)my_sheets.Item["Sheet1"];

        Dictionary<string, string> location_dictionary = new Dictionary<string, string>();
        public void InitializeLocations()
        {
            Excel.Range my_range = locations.Range["B2:B42"];
            Excel.Range my_range2 = locations.Range["D2:D42"];

            if (my_range != null && my_range2 != null)
            {
                for (int i = 1; i < locations.UsedRange.Rows.Count; i++)
                {
                    string key = ((Excel.Range)locations.Cells[i+1, 2]).Value2;
                    string value = ((Excel.Range)locations.Cells[i+1, 4]).Value2;

                    location_dictionary.Add(key, value);
                }
            }
        }
    }
}