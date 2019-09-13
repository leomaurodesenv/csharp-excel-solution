using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
// Excel
using System.Data;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.HPSF;
using NPOI.POIFS.FileSystem;
using NPOI.SS.UserModel;
using NPOI.HSSF.Util;
using NPOI.XSSF.UserModel;

namespace excelExample
{
    class Solution7
    {
        public static void run(DataTable table, string fileName, string extension = "xlsx")
        {
            // Workbook
            IWorkbook workbook = (extension == "xlsx") ? (IWorkbook)new XSSFWorkbook() : (IWorkbook)new HSSFWorkbook();

            // Check valid extension
            if (extension != "xls" && extension != "xlsx")
            {
                throw new Exception("This format is not supported");
            }
            
            // Process DataTable
            processDataTable(workbook, table);

            // Save file
            FileStream sw = File.Create(String.Format("{0}.{1}", fileName, extension));
            workbook.Write(sw);
            sw.Close();
        }

        private static void processDataTable(IWorkbook workbook, DataTable table)
        {
            // Variables
            IRow row;
            ISheet sheet = workbook.CreateSheet("sheet1");
            Dictionary<String, ICellStyle> styles = createStyles(workbook);

            // Make header
            row = sheet.CreateRow(0);
            for (int j = 0; j < table.Columns.Count; j++)
            {
                ICell cell = row.CreateCell(j);
                String columnName = table.Columns[j].ToString();
                cell.SetCellValue(columnName);
                cell.CellStyle = (styles["header"]);
            }

            // Add data row
            for (int i = 0; i < table.Rows.Count; i++)
            {
                row = sheet.CreateRow(i + 1);
                for (int j = 0; j < table.Columns.Count; j++)
                {
                    ICell cell = row.CreateCell(j);
                    String columnName = table.Columns[j].ToString();
                    cell.SetCellValue(table.Rows[i][columnName].ToString());
                    cell.CellStyle = (styles["normal"]);
                }
            }
        }

        private static Dictionary<String, ICellStyle> createStyles(IWorkbook workbook)
        {
            // Variables
            IFont font;
            ICellStyle style;
            Dictionary<String, ICellStyle> styles = new Dictionary<String, ICellStyle>();

            // Header cell
            font = workbook.CreateFont();
            font.Color = IndexedColors.White.Index;
            font.FontHeightInPoints = 9;
            font.IsBold = true;
            font.FontName = "Tahoma";
            style = CreateBorderedStyle(workbook);
            style.Alignment = HorizontalAlignment.Center;
            style.FillPattern = FillPattern.SolidForeground;
            style.FillForegroundColor = (IndexedColors.Green.Index);
            style.SetFont(font);
            styles.Add("header", style);

            // Normal cell
            font = workbook.CreateFont();
            font.Color = IndexedColors.Black.Index;
            font.FontHeightInPoints = 9;
            font.FontName = "Tahoma";
            style = CreateBorderedStyle(workbook);
            style.SetFont(font);
            styles.Add("normal", style);

            return styles;
        }

        private static ICellStyle CreateBorderedStyle(IWorkbook wb)
        {
            // Border entire cell
            ICellStyle style = wb.CreateCellStyle();
            style.BorderRight = BorderStyle.Thin;
            style.RightBorderColor = (IndexedColors.Black.Index);
            style.BorderBottom = BorderStyle.Thin;
            style.BottomBorderColor = (IndexedColors.Black.Index);
            style.BorderLeft = BorderStyle.Thin;
            style.LeftBorderColor = (IndexedColors.Black.Index);
            style.BorderTop = BorderStyle.Thin;
            style.TopBorderColor = (IndexedColors.Black.Index);
            return style;
        }
    }
}
