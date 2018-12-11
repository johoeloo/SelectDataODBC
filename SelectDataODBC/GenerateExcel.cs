using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Office.Interop.Excel;

namespace SelectDataODBC
{
    class GenerateExcel
    {
        public void GenerateExcelFile(string[] columns, int noOfColumns ,string[] data, string saveAs)
        {
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            app.Visible = true;
            app.WindowState = XlWindowState.xlMaximized;

            Workbook wb = app.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            Worksheet ws = wb.Worksheets[1];
            DateTime currentDate = DateTime.Now;
            int column = 0;
            int row = 1;
            for (int col = 0; col < noOfColumns; col++)
            {
                ws.Cells[col + 1][1].Value = columns[col];
            }

           for (int i = 0; i < data.Length; i++)
           {
                if (column> noOfColumns-1)
                    column = 0;

                if (column==0)
                    row = row + 1;
            
                ws.Cells[column + 1][row].Value = data[i];
                column = column + 1;
            }


            wb.SaveCopyAs(saveAs + "\\export"+ currentDate.ToShortDateString()+ ".xlsx");

        }
    }
}
