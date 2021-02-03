using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace MSPP_Report
{
    class WorkSheetSetUp
    {
        public void openFile(string[] args)
        {
            Application excelApp = new Application();
            Workbook excelBook = excelApp.Workbooks.Open(args[0]);
            _Worksheet excelSheet = excelBook.Sheets[1];
            Range excelRange = excelSheet.UsedRange;

            int row = excelRange.Rows.Count;
            int cols = excelRange.Columns.Count;

            excelSheet.PageSetup.Orientation = XlPageOrientation.xlLandscape;
            excelSheet.PageSetup.Zoom = 85;
            excelSheet.Columns["A:H"].ColumnWidth = 16.9;
            excelSheet.Columns["B:C"].HorizontalAlignment = XlHAlign.xlHAlignLeft;
            excelSheet.Cells[11, 4].HorizontalAligment = XlHAlign.xlHAlignRight;
            excelSheet.Range[excelSheet.Cells[11,4], excelSheet.Cells[11,8]].Font.Bold = true;

            Range findInColumn = excelSheet.Columns["A:A"];
            Range findInCell = findInColumn.Find("Customer Name");
            Range findNext = findInCell.FindNext(findInCell);

            int y = 0;
            int sumRange = 0;

            while (y < 2)
            {
                excelSheet.HPageBreaks.Add(findNext);
                if (y == 0)
                {
                    sumRange = findNext.Row - 8;
                }
                findInCell = findNext;
                findNext = findInColumn.FindNext(findInCell);

                y++;
            }
        }
    }
}
