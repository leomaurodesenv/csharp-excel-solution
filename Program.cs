using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
// Excel
using System.Data;

namespace excelExample
{
    class Program
    {
        static void Main(string[] args)
        {
            String extesion = "xlsx";
            String fileName = "fileExample";
            DataTable table = GetDataTable();

            // - Solutions:
            // 1. ClosedXML - Not support excel 2003 (xls)
            // 2. Microsoft.Office.Interop.Excel - Not recommend for Server, also is pay
            // 3. GemBox.Spreadsheet - Pay
            // 4. EASY XLS - Pay
            // 5. EPPlus - Only xlsx
            // 6. ExcelLibrary
            // 7. NPOI - Support xlsx, xls

            //Solution1.run(table, fileName, extesion);
            //Solution5.run(table, fileName, extesion);
            //Solution6.run(table, fileName, extesion);
            Solution7.run(table, fileName, extesion);

            Console.WriteLine("Press to continue...");
            Console.ReadKey();
        }

        private static DataTable GetDataTable ()
        {
            //Create a DataTable with four columns
            DataTable table = new DataTable();
            table.Columns.Add("Id", typeof(int));
            table.Columns.Add("Type", typeof(string));
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Date", typeof(DateTime));

            //Add five DataRows
            table.Rows.Add(25, "Indocin", "David", DateTime.Now);
            table.Rows.Add(50, "Enebrel", "Sam", DateTime.Now);
            table.Rows.Add(10, "-H3", "Christoff", DateTime.Now);
            table.Rows.Add(21, "Combivent", "Janet", DateTime.Now);
            table.Rows.Add(100, "Dilantin", "-H4", "10/09/2019  18:25:58");

            return table;
        }
    }
}
