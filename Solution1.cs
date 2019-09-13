using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
// Excel
using System.Data;
using ClosedXML.Excel;

namespace excelExample
{
    class Solution1
    {
        public static void run (DataTable table, String fileName)
        {
            XLWorkbook worker = new XLWorkbook();
            worker.Worksheets.Add(table, "WorksheetName");
            worker.SaveAs(fileName); // support only xlsx
        }
    }
}
