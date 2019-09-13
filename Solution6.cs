using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
// Excel
using System.Data;
using ExcelLibrary.SpreadSheet;
using System.IO;

namespace excelExample
{
    class Solution6
    {
        public static void run(DataTable table, string fileName, string extension = "xls")
        {
            if (File.Exists(fileName))
            {
                File.Delete(fileName);
            }

            Workbook workbook = new Workbook();
            workbook.Worksheets.Clear();

            Worksheet worksheet = new Worksheet("sheet1");
            processSheet(table, worksheet);
            workbook.Worksheets.Add(worksheet);

            workbook.Save(String.Format("{0}.{1}", fileName, extension)); // xlsx not supported
        }

        private static void processSheet(DataTable table, Worksheet ws)
        {
            // Make header
            foreach (DataColumn dataColumn in table.Columns)
            {
                var headerCell = new Cell(string.IsNullOrEmpty(dataColumn.Caption)
                                              ? dataColumn.ColumnName
                                              : dataColumn.Caption,
                                          CellFormat.General);
                ws.Cells[0, dataColumn.Ordinal] = headerCell;
            }

            // Make body
            var rowIndex = 1;
            foreach (DataRow sourceRow in table.Rows)
            {
                foreach (DataColumn sourceColumn in table.Columns)
                {
                    var sourceValue = sourceRow[sourceColumn];
                    if (sourceValue == DBNull.Value) sourceValue = null;

                    var destinationCell = new Cell(sourceValue);
                    ws.Cells[rowIndex, sourceColumn.Ordinal] = destinationCell;
                }
                rowIndex++;
            }
        }
    }
}
