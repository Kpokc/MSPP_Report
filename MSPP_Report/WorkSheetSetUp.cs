using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Diagnostics;

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
            excelSheet.Range[excelSheet.Cells[11, 4], excelSheet.Cells[11, 8]].Font.Bold = true;


            /// Adds Pagebreaks at "Customer Name:" position //////
            Range findInColumn = excelSheet.Columns["A:A"];
            Range findInCell = findInColumn.Find("Customer Name:");
            Range findNext = findInCell.FindNext(findInCell);

            for (int custNamePosition = 0; custNamePosition < 3; custNamePosition++)
            {
                findInCell = findNext;
                findNext = findInColumn.FindNext(findInCell);
                if (custNamePosition < 2)
                {
                    excelSheet.HPageBreaks.Add(findNext);
                }
                else
                {
                    Console.WriteLine(findNext.Row);
                    break;
                }
            }
            /// ////////////////////////////////// //////

            ///////// Create new file name and Save as to ////////
            DateTime today = DateTime.Now;

            string newFileName = @"V:\Warehouses\Parkmore Warehouse\Reports\Medtronic Reports\Spare Parts (MSPP) SSL B1\MSPP Report " + today.ToString("yyyyMMdd") + ".xls";

            excelBook.SaveAs(newFileName, XlFileFormat.xlWorkbookNormal);

            /////// Pass new file location to SaveToEmailDrafts Class
            SaveToEmailDrafts saveToEmail = new SaveToEmailDrafts();
            saveToEmail.addFileToEmail(newFileName, today.ToString("yyyyMMdd"));
            //////////////////////////////////////////////////////

            Console.WriteLine("MSPP Report " + today.ToString("yyyyMMdd") + "  ----> Saved");
            
            excelApp.Quit();


            //////////// Quit and Release /////////////
            Marshal.ReleaseComObject(excelApp);

            GC.Collect();
            GC.WaitForPendingFinalizers();

            Process[] List;
            List = Process.GetProcessesByName("EXCEL");
            foreach (Process proc in List)
            {
                proc.Kill();
            }
        }
    }
}
