using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace CBD
{
    class Algorithm
    {
        private List<Bus> buses = new List<Bus>();
        private List<Group> groups = new List<Group>();

        //Create COM Objects for everything that is referenced
        public static Excel.Application my_excel = new Excel.Application();
        public static Excel.Workbook my_book = my_excel.Workbooks.Open(Environment.CurrentDirectory + "\\Locations.xlsx");
        public static Excel._Worksheet my_groups = my_book.Sheets[1];
        public static Excel._Worksheet my_buses = my_book.Sheets[2];


        public void InitializeGroups() {
            Excel.Range used_range = my_groups.UsedRange;
            int[] columns = { 3, 4, 5 };
            foreach (Excel.Range row in used_range.Rows) {
                string name = my_groups.Cells[row.Row, columns[0]].Value2.ToString() + my_groups.Cells[row.Row, columns[1]].Value2.ToString();
                int size = my_groups.Cells[row.Row, columns[2]].Value2;
                Group g = new Group(name, size);
                groups.Add(g);
            }
        }
        public void InitializeBuses() {
            Excel.Range used_range = my_buses.UsedRange;
            int[] columns = { 1, 2 };
            foreach(Excel.Range row in used_range.Rows) {
                int num = my_buses.Cells[row.Row, columns[0]].Value2;
                int cap = my_buses.Cells[row.Row, columns[1]].Value2;
                Bus b = new Bus(num, cap);
                buses.Add(b);
            }
        }
        private void CloseExcel()
        {
            my_book.Close(true, null, null);
            my_excel.Quit();
        }
    }
}
