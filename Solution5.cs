using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
// Excel
using System.Data;
using OfficeOpenXml;
using System.IO;

namespace excelExample
{
    class Solution5
    {
        public static void run (DataTable table, String fileName)
        {
            using (ExcelPackage worker = new ExcelPackage())
            {
                ExcelWorksheet worksheet = worker.Workbook.Worksheets.Add("sheet1");
                // Add table data
                worksheet.Cells["A1"].LoadFromDataTable(table, true);
                FileInfo file = new FileInfo(fileName);
                worker.SaveAs(file); // support only xlsx
            }

            
        }
    }
}
