using System;
using System.CodeDom;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace excel
{
    class Excel
    {
        static _Application excel = new _Excel.Application();
        static Workbook wb;
        static Worksheet ws;

        public static bool SaveFile(string path, string fileName,Control panel)
        {
            if (excel == null)
                return false;

            wb = excel.Workbooks.Add();
            ws = (_Excel.Worksheet)wb.Worksheets.Item[1];

            //ws.Cells[1, 1] = "Hi";
            foreach (System.Windows.Forms.TextBox txtBox in panel.Controls)
            {
                WriteCells(ws, txtBox.Name, txtBox.Text);
            }

            string format = fileName + "{0}";
            int n = 1;

            while (File.Exists(path + @"\" + fileName + @".xls"))
            {
                fileName = string.Format(format, "(" + (n++) + ")");
            }
            
            wb.SaveAs(path + @"\" + fileName + @".xls", _Excel.XlFileFormat.xlWorkbookNormal);
            wb.Close(true);
            excel.Quit();

            /*// use when you wanna close programe for deallocate somthing ??
            Marshal.ReleaseComObject(ws);
            Marshal.ReleaseComObject(wb);
            Marshal.ReleaseComObject(excel);
            */
            return true;

        }

     public static bool WriteCells(Worksheet ws, string boxName, string data)
        {
            char[] spearator = {'_'};
            string[] rowCol = boxName.Split(spearator);
            int col = int.Parse(rowCol[0]);
            int row = int.Parse(rowCol[1]);
            ws.Cells[row, col] = data;
            return true;
        }

    }
}
