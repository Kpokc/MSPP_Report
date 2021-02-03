using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Microsoft.Office.Interop.Excel;

namespace MSPP_Report
{
    class Program
    {
        static void Main(string[] args)
        {
            ///Checking if excel is installed//////////////////

            Application excelApp = new Application();

            if (excelApp == null)
            {
                Console.WriteLine("Excel is not installed.");
            }

            //////////////////////////////////////////////////

            /// If file exists pass it to next class ////////////
            string path = @"C:\Users\ssladmin\Desktop\Weekly rep\";
            bool fileExist = File.Exists(path);
            if (!fileExist)
            {   
                ///////// Get all files links /////////////////
                string[] locationArray = Directory.GetFiles(path);
                locationArray = Array.ConvertAll(locationArray, x => x.ToUpper());

                /// Array with locations passed to WorksheetSetUp class ////
                WorkSheetSetUp passArray = new WorkSheetSetUp();
                passArray.openFile(locationArray);

                /////// Delete all files in folder ///////////////
                Console.WriteLine("\nPres Enter to Delete all files and close console...");
                Console.ReadLine();
                foreach (string fileLink in locationArray)
                {
                    /*File.Delete(fileLink);*/
                }
            }   
            else 
            {
                Console.WriteLine("There is no files in the folder.");
                Console.ReadLine();
            }
        }
    }
}
