using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;       //microsoft Excel 14 object in references-> COM tab

namespace Sandbox
{
    public class Read_From_Excel
    {
        static void Main(string[] args)
        {
            string conFile = @"\\grc652\MedPhysics Backup\Data Extractions\Automation\Template Creation Input\Brain\SRS_2003-2019.xls";

            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApplPat = new Excel.Application();
            Excel.Workbook xlWorkbooksPat = xlApplPat.Workbooks.Open(conFile);
            Excel._Worksheet xlWorksheetsPat = xlWorkbooksPat.Sheets[1];
            Excel.Range xlRangePat = xlWorksheetsPat.UsedRange;

            int rowCountPat = xlRangePat.Rows.Count;
            int colCountPat = xlRangePat.Columns.Count;

            //iterate over the rows and columns and print to the console as it appears in the file
            //excel is not zero based!!
            for (int i = 2; i <= rowCountPat; i++)
            {
                string cell;
                cell = xlRangePat.Cells[i, 1].Value2.ToString();
                if (cell.Length == 7)
                {
                    string cellNew = $"0{cell}";
                    xlRangePat.Cells[i, 1].Value = cellNew;
                }
                else if (cell.Length == 6)
                {
                    string cellNew = $"00{cell}";
                    xlRangePat.Cells[i, 1].Value = cellNew;
                }
                else if (cell.Length == 5)
                {
                    string cellNew = $"000{cell}";
                    xlRangePat.Cells[i, 1].Value = cellNew;
                }
                else if (cell.Length == 4)
                {
                    string cellNew = $"0000{cell}";
                    xlRangePat.Cells[i, 1].Value = cellNew;
                }
            }

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(xlRangePat);
            Marshal.ReleaseComObject(xlWorksheetsPat);

            //close and release
            xlWorkbooksPat.Close();
            Marshal.ReleaseComObject(xlWorkbooksPat);

            //quit and release
            xlApplPat.Quit();
            Marshal.ReleaseComObject(xlApplPat);
        }
    }
}