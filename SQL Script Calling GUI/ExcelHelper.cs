using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Forms;
using VMS.TPS.Common.Model.API;
using Excel = Microsoft.Office.Interop.Excel;
using MessageBox = System.Windows.Forms.MessageBox;
using MessageBoxOptions1 = System.Windows.Forms.MessageBoxOptions;


namespace ExcelExtensions
{
    class ExcelHelper
    {
        // Private Fields
        Excel.Application xlApp;
        Excel.Workbook xlBook;
        Excel.Sheets xlSheets;
        Excel.Worksheet xlSheet;

        Patient patient;
        Dictionary<string, string> excelIDs;

        string SaveDirectory { get; set; }
        string FileNameWithExt { get; set; }
        string FilePathWithExt { get; set; }


        // Constructor
        public ExcelHelper(Patient patient = null, string filePath = null)
        {
            xlApp = null;
            xlBook = null;
            xlSheets = null;
            xlSheet = null;

            this.patient = patient;

            this.excelIDs = new Dictionary<string, string>();

            if (!string.IsNullOrWhiteSpace(filePath))
            {
                InitializeFileNamesAndPaths(filePath);
            }
        }

        // Accessors
        public void SetFilePath(string filePath)
        {
            if (!string.IsNullOrWhiteSpace(filePath))
            {
                InitializeFileNamesAndPaths(filePath);
            }
        }

        public void AddToExcelIDs(string eclipsePlanID, string excelPlanName)
        {
            eclipsePlanID = eclipsePlanID.ToUpper();

            if (!excelIDs.ContainsKey(eclipsePlanID))
            {
                excelIDs.Add(eclipsePlanID, excelPlanName);
            }
        }

        // Public Methods
        // Checks if Excel is installed on the computer, and exits the program if it isn't.
        public void CheckInstallation()
        {
            uint timeout = 0; //messagebox timeout timer
            Type officeType = Type.GetTypeFromProgID("Excel.Application");
            if (officeType == null)
            {
                //Self-exiting messagebox pop up (change timeout timer to change exit time)
                MessageBoxEx.Show($"Microsoft Excel is not installed on this computer!", "Extraction Failed",
                                MessageBoxButtons.OK, MessageBoxIcon.Error, timeout);
                System.Environment.Exit(0);
            }
        }

        // Initializes the file names and file paths (with and without the .xlxs extension).
        public void InitializeFileNamesAndPaths(string filePath)
        {
            
            uint timeout = 0; //messagebox timeout timer
            SaveDirectory = Path.GetDirectoryName(filePath);
            // Add a backslash character ('\') to the end of the save directory path.
            if (!SaveDirectory.EndsWith(@"\"))
            {
                SaveDirectory = string.Format(@"{0}\", SaveDirectory);
            }

            string excelFileName = Path.GetFileName(filePath);

            //FileNameNoExt = Path.ChangeExtension(excelFileName, null);

            // Add the extension .xlxs to the end of the filename.
            string fileExtension = Path.GetExtension(excelFileName);

            if (fileExtension == null)
            {
                //Self-exiting messagebox pop up (change timeout timer to change exit time)
                MessageBoxEx.Show($"The file path could not be found.", "Extraction Failed",
                                            MessageBoxButtons.OK, MessageBoxIcon.Error, timeout);
                return;
            }
            else if (fileExtension == "")
            {
                FileNameWithExt = string.Format(@"{0}.xlsx", excelFileName);
            }
            else if (fileExtension != ".xlsx")
            {
                FileNameWithExt = Path.ChangeExtension(excelFileName, ".xlsx");
            }
            else
            {
                FileNameWithExt = excelFileName;
            }

            FilePathWithExt = $"{SaveDirectory}{FileNameWithExt}";
            //FilePathNoExt = $"{saveDirectory}{FileNameNoExt}";
        }

        public void CreateDirectory()
        {
            Directory.CreateDirectory(SaveDirectory);
        }

        public void InitializeExcelObjects()
        {
            object missing = System.Reflection.Missing.Value;

            // Used to get all running Excel instances from processes.
            ExcelAppCollection eac = new ExcelAppCollection();

            // Get currently running Excel processes.
            System.Diagnostics.Process[] excelProcesses = System.Diagnostics.Process.GetProcessesByName("EXCEL");

            // Check if an Excel process is already running.
            if (!IsNullOrEmpty(excelProcesses))
            {
                foreach (System.Diagnostics.Process process in excelProcesses)
                {
                    // Get a reference to an instance of an Excel application.
                    try
                    {
                        xlApp = eac.FromProcess(process);
                    }
                    catch (Exception e)
                    {
                        // This is the case where an Excel process is running but
                        // no Excel Application is open.
                        System.Diagnostics.Debug.WriteLine($"Exception: {e}");
                    }

                    if (xlApp == null)
                    {
                        continue;
                    }

                    xlApp.DisplayAlerts = false;

                    // Check all workbooks in the Excel instance to see if the file
                    // we want is already open.
                    foreach (Excel.Workbook workbook in xlApp.Workbooks)
                    {
                        if (workbook.Name == FileNameWithExt)
                        {
                            xlBook = workbook;
                            break;
                        }
                    }

                    if (xlBook != null)
                    {
                        break;
                    }
                }
            }

            if (xlBook == null)
            {
                // Create a new Excel application.
                xlApp = new Excel.Application
                {
                    SheetsInNewWorkbook = 1,
                    DisplayAlerts = false,
                    Visible = true
                };

                if (File.Exists(FilePathWithExt))
                {
                    xlBook = xlApp.Workbooks.Open(FilePathWithExt);
                }
                else
                {
                    xlBook = xlApp.Workbooks.Add(missing);
                }
            }

            xlSheets = xlBook.Worksheets;
        }

        // Returns a bool depending on whether a new worksheet was created or not.
        public bool GetOrCreateWorksheet(string sheetName)
        {
            object missing = System.Reflection.Missing.Value;

            // Get existing worksheet with plan name if it exists, else create it.
            if (WorksheetExists(this.xlSheets, sheetName))
            {
                xlSheet = GetWorksheet(this.xlSheets, sheetName);
                xlSheet.Select(missing);
                return false;
            }
            else
            {
                xlSheet = CreateWorksheet(this.xlSheets, sheetName);
                return true;
            }
        }

        public void ScrollRowIntoView(int row)
        {
            if (xlApp == null)
            {
                return;
            }

            Excel.Range visibleRange = xlApp.ActiveWindow.VisibleRange;
            // Check if row we are writing to is visible.
            if (row < visibleRange.Row || row > (visibleRange.Row + visibleRange.Rows.Count - 2))
            {
                // Scroll Excel worksheet if the row we are writing to is offscreen.
                int rowToScrollTo = row - visibleRange.Rows.Count;

                if (rowToScrollTo < 1)
                {
                    xlApp.ActiveWindow.ScrollRow = 1;
                }
                else
                {
                    xlApp.ActiveWindow.ScrollRow = rowToScrollTo;
                }
            }
        }

        public void SaveWorkbook()
        {
            if (xlApp == null || xlBook == null)
            {
                return;
            }

            object missing = System.Reflection.Missing.Value;

            if (File.Exists(FilePathWithExt))
            {
                xlBook.Save();
            }
            else
            {
                try
                {
                    string filePathNoExt = Path.ChangeExtension(FilePathWithExt, null);

                    xlBook.SaveAs(filePathNoExt, Excel.XlFileFormat.xlOpenXMLWorkbook, missing,
                    missing, false, true, Excel.XlSaveAsAccessMode.xlExclusive,
                    Excel.XlSaveConflictResolution.xlUserResolution, true,
                    missing, missing, missing);
                }
                catch (Exception e)
                {
                    System.Diagnostics.Debug.WriteLine($"Exception: '{e}'");
                }
            }
        }

        public void ExportDoseAsExcel(List<PlanQueries> queryList)
        {
            object missing = System.Reflection.Missing.Value;

            CreateDirectory();

            string studyID = patient.GenerateStudyID();

            InitializeExcelObjects();

            RemoveExcelHighlights(xlBook);

            // Loop through the list of lists of queries in the same plan.
            foreach (PlanQueries queriesByPlan in queryList)
            {
                // Get the Excel plan name of the first query in the plan list.
                string planName = excelIDs[queriesByPlan.GetPlanID().ToUpper()];

                // Used for helping to check whether to create headers/subheaders.
                bool newWorksheet = GetOrCreateWorksheet(planName);

                int rowToWrite = GetPatientRowOrBlankRow(xlSheet, studyID, 2, queriesByPlan.GetPlanID(), 6);

                if (rowToWrite == -1)
                {
                    continue;
                }

                ScrollRowIntoView(rowToWrite);

                int rowMainIndex = 1;
                int rowSubIndex = 2;
                int colIndex = 2;

                if (newWorksheet)
                {
                    if (WorksheetExists(xlSheets, "Sheet1"))
                    {
                        xlSheets["Sheet1"].Delete();
                    }

                    // Set up generic patient headers and subheaders.
                    xlSheet.Cells[rowMainIndex, colIndex] = "Patient Information";
                    xlSheet.Cells[rowSubIndex, colIndex] = "Study ID";
                    colIndex++;
                    xlSheet.Cells[rowMainIndex, colIndex] = "Patient Information";
                    xlSheet.Cells[rowSubIndex, colIndex] = "Date of Birth";
                    colIndex++;
                    xlSheet.Cells[rowMainIndex, colIndex] = "Patient Information";
                    xlSheet.Cells[rowSubIndex, colIndex] = "Age at Date of Plan Creation";
                    colIndex++;
                    xlSheet.Cells[rowMainIndex, colIndex] = "Plan Information";
                    xlSheet.Cells[rowSubIndex, colIndex] = "Date of Plan Creation";
                    colIndex++;
                    xlSheet.Cells[rowMainIndex, colIndex] = "Plan Information";
                    xlSheet.Cells[rowSubIndex, colIndex] = "Plan ID";
                    colIndex++;

                    // If Plan Sum.
                    if (queriesByPlan.IsPlanSum())
                    {
                        int planCount = 1;

                        foreach (PlanQueries plan in queriesByPlan.GetPlanSumPlans())
                        {
                            xlSheet.Cells[rowMainIndex, colIndex] = $"Plan {planCount} Prescription";
                            xlSheet.Cells[rowSubIndex, colIndex] = $"Plan {planCount} ID";
                            colIndex++;

                            xlSheet.Cells[rowMainIndex, colIndex] = $"Plan {planCount} Prescription";
                            xlSheet.Cells[rowSubIndex, colIndex] = "Dose (Gy)/fx";
                            colIndex++;

                            xlSheet.Cells[rowMainIndex, colIndex] = $"Plan {planCount} Prescription";
                            xlSheet.Cells[rowSubIndex, colIndex] = "Number of Fractions";
                            colIndex++;

                            xlSheet.Cells[rowMainIndex, colIndex] = $"Plan {planCount} Prescription";
                            xlSheet.Cells[rowSubIndex, colIndex] = "Total Dose (Gy)";
                            colIndex++;

                            xlSheet.Cells[rowMainIndex, colIndex] = $"Plan {planCount} Field Information";
                            xlSheet.Cells[rowSubIndex, colIndex] = "Number of Fields (Beams)";
                            colIndex++;

                            xlSheet.Cells[rowMainIndex, colIndex] = $"Plan {planCount} Field Information";
                            xlSheet.Cells[rowSubIndex, colIndex] = "Plan Type";
                            colIndex++;

                            xlSheet.Cells[rowMainIndex, colIndex] = $"Plan {planCount} Field Information";
                            xlSheet.Cells[rowSubIndex, colIndex] = "Energy Mode";
                            colIndex++;

                            xlSheet.Cells[rowMainIndex, colIndex] = $"Plan {planCount} Field Information";
                            xlSheet.Cells[rowSubIndex, colIndex] = "SSD 1 (cm)";
                            colIndex++;

                            xlSheet.Cells[rowMainIndex, colIndex] = $"Plan {planCount} Field Information";
                            xlSheet.Cells[rowSubIndex, colIndex] = "SSD 2 (cm)";
                            colIndex++;

                            xlSheet.Cells[rowMainIndex, colIndex] = $"Plan {planCount} Field Information";
                            xlSheet.Cells[rowSubIndex, colIndex] = "SSD Separation (cm)";
                            colIndex++;

                            planCount++;
                        }
                    }
                    // If Plan Setup.
                    else
                    {
                        xlSheet.Cells[rowMainIndex, colIndex] = "Prescription";
                        xlSheet.Cells[rowSubIndex, colIndex] = "Dose (Gy)/fx";
                        colIndex++;

                        xlSheet.Cells[rowMainIndex, colIndex] = $"Prescription";
                        xlSheet.Cells[rowSubIndex, colIndex] = "Number of Fractions";
                        colIndex++;

                        xlSheet.Cells[rowMainIndex, colIndex] = $"Prescription";
                        xlSheet.Cells[rowSubIndex, colIndex] = "Total Dose (Gy)";
                        colIndex++;

                        xlSheet.Cells[rowMainIndex, colIndex] = $"Field Information";
                        xlSheet.Cells[rowSubIndex, colIndex] = "Number of Fields (Beams)";
                        colIndex++;

                        xlSheet.Cells[rowMainIndex, colIndex] = $"Field Information";
                        xlSheet.Cells[rowSubIndex, colIndex] = "Plan Type";
                        colIndex++;

                        xlSheet.Cells[rowMainIndex, colIndex] = $"Field Information";
                        xlSheet.Cells[rowSubIndex, colIndex] = "Energy Mode";
                        colIndex++;

                        xlSheet.Cells[rowMainIndex, colIndex] = $"Field Information";
                        xlSheet.Cells[rowSubIndex, colIndex] = "SSD 1 (cm)";
                        colIndex++;

                        xlSheet.Cells[rowMainIndex, colIndex] = $"Field Information";
                        xlSheet.Cells[rowSubIndex, colIndex] = "SSD 2 (cm)";
                        colIndex++;

                        xlSheet.Cells[rowMainIndex, colIndex] = $"Field Information";
                        xlSheet.Cells[rowSubIndex, colIndex] = "SSD Separation (cm)";
                        colIndex++;
                    }


                    for (int i = 2; i < colIndex; i++)
                    {
                        FormatMainHeaderCell(xlSheet.Cells[rowMainIndex, i]);
                        FormatSubHeaderCell(xlSheet.Cells[rowSubIndex, i]);
                    }
                }

                colIndex = 2;

                // Fill in generic patient information.
                if (string.IsNullOrWhiteSpace(Convert.ToString(xlSheet.Cells[rowToWrite, 1].Value2)))
                {
                    xlSheet.Cells[rowToWrite, 1] = rowToWrite - 2;
                    xlSheet.Cells[rowToWrite, 1].Font.Bold = true;
                }
                xlSheet.Cells[rowToWrite, colIndex] = studyID;
                colIndex++;
                xlSheet.Cells[rowToWrite, colIndex] = patient.DateOfBirth;
                colIndex++;
                xlSheet.Cells[rowToWrite, colIndex] = queriesByPlan.GetAgeAtPlanCreationDate();
                colIndex++;

                if (queriesByPlan.GetPlanCreationDateTime() is DateTime dt)
                {
                    xlSheet.Cells[rowToWrite, colIndex] = dt.Date;
                }
                else
                {
                    xlSheet.Cells[rowToWrite, colIndex] = null;
                }
                colIndex++;

                xlSheet.Cells[rowToWrite, colIndex] = queriesByPlan.GetPlanID();
                colIndex++;

                // If Plan Sum.
                if (queriesByPlan.IsPlanSum())
                {
                    foreach (PlanQueries plan in queriesByPlan.GetPlanSumPlans())
                    {
                        xlSheet.Cells[rowToWrite, colIndex] = plan.GetPlanID();
                        colIndex++;
                        xlSheet.Cells[rowToWrite, colIndex] = plan.GetDoseFx();
                        colIndex++;
                        xlSheet.Cells[rowToWrite, colIndex] = plan.GetNumOfFractions();
                        colIndex++;
                        xlSheet.Cells[rowToWrite, colIndex] = plan.GetTotalDoseInGy();
                        colIndex++;
                        xlSheet.Cells[rowToWrite, colIndex] = plan.GetNumOfBeams();
                        colIndex++;
                        xlSheet.Cells[rowToWrite, colIndex] = plan.GetMlcPlanType();
                        colIndex++;
                        xlSheet.Cells[rowToWrite, colIndex] = plan.GetFieldEnergies();
                        colIndex++;
                        xlSheet.Cells[rowToWrite, colIndex] = plan.GetSsd1();
                        xlSheet.Cells[rowToWrite, colIndex].NumberFormat = "0.0##;;0";
                        colIndex++;
                        xlSheet.Cells[rowToWrite, colIndex] = plan.GetSsd2();
                        xlSheet.Cells[rowToWrite, colIndex].NumberFormat = "0.0##;;0";
                        colIndex++;
                        xlSheet.Cells[rowToWrite, colIndex] = plan.GetSsdDivision();
                        xlSheet.Cells[rowToWrite, colIndex].NumberFormat = "0.0##;;0";
                        colIndex++;
                    }
                }
                // If Plan Setup.
                else
                {
                    xlSheet.Cells[rowToWrite, colIndex] = queriesByPlan.GetDoseFx();
                    colIndex++;
                    xlSheet.Cells[rowToWrite, colIndex] = queriesByPlan.GetNumOfFractions();
                    colIndex++;
                    xlSheet.Cells[rowToWrite, colIndex] = queriesByPlan.GetTotalDoseInGy();
                    colIndex++;
                    xlSheet.Cells[rowToWrite, colIndex] = queriesByPlan.GetNumOfBeams();
                    colIndex++;
                    xlSheet.Cells[rowToWrite, colIndex] = queriesByPlan.GetMlcPlanType();
                    colIndex++;
                    xlSheet.Cells[rowToWrite, colIndex] = queriesByPlan.GetFieldEnergies();
                    colIndex++;
                    xlSheet.Cells[rowToWrite, colIndex] = queriesByPlan.GetSsd1();
                    xlSheet.Cells[rowToWrite, colIndex].NumberFormat = "0.0##;;0";
                    colIndex++;
                    xlSheet.Cells[rowToWrite, colIndex] = queriesByPlan.GetSsd2();
                    xlSheet.Cells[rowToWrite, colIndex].NumberFormat = "0.0##;;0";
                    colIndex++;
                    xlSheet.Cells[rowToWrite, colIndex] = queriesByPlan.GetSsdDivision();
                    xlSheet.Cells[rowToWrite, colIndex].NumberFormat = "0.0##;;0";
                    colIndex++;
                }

                for (int i = 1; i < colIndex; i++)
                {
                    FormatBodyCell(xlSheet.Cells[rowToWrite, i]);
                    xlSheet.Cells[rowToWrite, i].Interior.Color = System.Drawing.Color.Yellow;
                }

                int numOfQueries = queriesByPlan.GetDVHQueryCount();

                // Loop through the queries with the same plan. These will be on the same worksheet.
                for (int queryIndex = 0; queryIndex < numOfQueries; queryIndex++)
                {
                    // Write to header/subheader cells if they are empty.
                    if (newWorksheet)
                    {
                        xlSheet.Cells[1, colIndex] = queriesByPlan.GetDVHQueryList()[queryIndex].GetStructureName();
                        FormatMainHeaderCell(xlSheet.Cells[1, colIndex]);
                        xlSheet.Cells[2, colIndex] = queriesByPlan.GetDVHQueryList()[queryIndex].GetQueryString();
                        FormatSubHeaderCell(xlSheet.Cells[2, colIndex]);

                        // Last query, so we merge the header cells.
                        if (queryIndex + 1 == numOfQueries)
                        {
                            MergeAdjacentCells(xlSheet, 1);
                            xlSheet.Rows[2].RowHeight = xlSheet.Rows[1].RowHeight * 4;
                            xlSheet.Columns.ColumnWidth = xlSheet.Columns[1].ColumnWidth * 1.2;
                            xlSheet.Columns[1].ColumnWidth = xlSheet.Columns[1].ColumnWidth * 0.5;
                            xlSheet.Rows[2].WrapText = true;
                        }
                    }

                    // Overwrite patient row or write to first blank row.
                    var cellValue = queriesByPlan.GetDVHQueryList()[queryIndex].GetDVHValue();
                    if (Double.IsNaN(cellValue))
                    {
                        xlSheet.Cells[rowToWrite, colIndex] = "";
                    }
                    else
                    {
                        string metric = queriesByPlan.GetDVHQueryList()[queryIndex].GetDVHMetric();
                        if (metric == "%")
                        {
                            xlSheet.Cells[rowToWrite, colIndex] = cellValue / 100;
                            xlSheet.Cells[rowToWrite, colIndex].NumberFormat = "0.0# %;;0 %";
                        }
                        else
                        {
                            xlSheet.Cells[rowToWrite, colIndex] = cellValue;
                            xlSheet.Cells[rowToWrite, colIndex].NumberFormat = "0.0##;;0";
                        }
                    }

                    FormatBodyCell(xlSheet.Cells[rowToWrite, colIndex]);
                    xlSheet.Cells[rowToWrite, colIndex].Interior.Color = System.Drawing.Color.Yellow;

                    colIndex++;
                }
            }

            ReleaseObject(xlSheet);

            xlApp.DisplayAlerts = true;

            SaveWorkbook();

            //xlBook.Close(0);

            ReleaseObject(xlSheets);
            ReleaseObject(xlBook);

            try
            {
                //xlApp.Quit();
                ReleaseObject(xlApp);
            }
            catch (InvalidComObjectException e)
            {
                System.Diagnostics.Debug.WriteLine($"Invalid Com object: '{e}'");
            }
        }

        private static void ReleaseObject(object obj)
        {
            if (obj == null)
            {
                return;
            }

            try
            {
                Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                System.Diagnostics.Debug.WriteLine("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        private void FormatMainHeaderCell(Excel.Range cells)
        {
            cells.Font.Bold = true;
            cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium);
        }

        private void FormatSubHeaderCell(Excel.Range cells)
        {
            cells.Font.Bold = true;
            cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin);
        }

        private void FormatBodyCell(Excel.Range cells)
        {
            cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin);
        }

        // Merges all adjacent cells with the same value in a specific row.
        private void MergeAdjacentCells(Excel.Worksheet sheet, int excelRow)
        {
            if (sheet == null || excelRow < 1)
            {
                return;
            }

            object missing = System.Reflection.Missing.Value;

            int lastColumn = GetFirstBlankExcelColumn(sheet);

            int columnIndex = 1;
            int cellsToMerge = 0;

            for (; columnIndex + 1 < lastColumn; columnIndex++)
            {
                string cellValue = sheet.Cells[excelRow, columnIndex].Value2;
                string cellValueNext = sheet.Cells[excelRow, columnIndex + 1].Value2;

                if (!string.IsNullOrWhiteSpace(cellValue))
                {
                    if (string.IsNullOrWhiteSpace(cellValueNext) || cellValue != cellValueNext)
                    {
                        if (cellsToMerge > 0)
                        {
                            sheet.Range[sheet.Cells[excelRow, columnIndex - cellsToMerge],
                                    sheet.Cells[excelRow, columnIndex]].Merge();

                            cellsToMerge = 0;
                        }
                    }
                    else if (!string.IsNullOrWhiteSpace(cellValueNext) && cellValue == cellValueNext)
                    {
                        cellsToMerge++;
                    }
                }
            }

            // Merge last group of cells with the same value
            sheet.Range[sheet.Cells[excelRow, columnIndex - cellsToMerge],
                        sheet.Cells[excelRow, columnIndex]].Merge();
        }

        // Returns the row number of the patient's plan in the worksheet if the entry already exists, and the
        // row number of the first blank row otherwise.
        private int GetPatientRowOrBlankRow(Excel.Worksheet sheet, string studyID, int studyIDColumn, string planID, int planIDColumn)
        {
            uint timeout = 0; //messagebox timeout timer
            if (sheet == null)
            {
                return -1;
            }

            object missing = System.Reflection.Missing.Value;

            int lastRow = Math.Max(3, GetFirstBlankExcelRow(sheet));

            Excel.Range pIDRange;

            //pIDRange = sheet.Cells[sheet.Cells[3, patientIDColumn], sheet.Cells[lastRow, patientIDColumn]];
            pIDRange = sheet.Columns[studyIDColumn];

            Excel.Range currentFind = null;
            Excel.Range firstFind = null;

            currentFind = pIDRange.Find(studyID, missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart,
                                        Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false,
                                        missing, missing);

            while (currentFind != null)
            {
                // Keep track of the first range you find. 
                if (firstFind == null)
                {
                    firstFind = currentFind;
                }

                // If you didn't move to a new range, you are done.
                else if (currentFind.get_Address(Excel.XlReferenceStyle.xlA1)
                        == firstFind.get_Address(Excel.XlReferenceStyle.xlA1))
                {
                    break;
                }

                Excel.Range planIdCell = sheet.Cells[currentFind.Row, planIDColumn];

                if (planIdCell.Value == planID)
                {
                    //Self-exiting messagebox pop up (change timeout timer to change exit time)
                    DialogResult dialogResult = MessageBoxEx.Show(new Form { TopMost = true }, $"This will overwrite Excel row {currentFind.Row} for the worksheet {sheet.Name}.", "Warning",
                                                                MessageBoxButtons.OKCancel, MessageBoxIcon.Warning, timeout);
                    if (dialogResult == DialogResult.Cancel)
                    {
                        return -1;
                    }
                    return currentFind.Row;
                }

                //currentFind.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                //currentFind.Font.Bold = true;

                currentFind = pIDRange.FindNext(currentFind);
            }

            return lastRow;
        }

        private int GetFirstBlankExcelRow(Excel.Worksheet sheet)
        {
            if (sheet == null)
            {
                return -1;
            }

            Excel.Range range = sheet.Cells[sheet.Rows.Count, 1];
            int lastRow = range.get_End(Excel.XlDirection.xlUp).Row;
            return lastRow + 1;
        }

        private int GetFirstBlankExcelColumn(Excel.Worksheet sheet)
        {
            if (sheet == null)
            {
                return -1;
            }

            Excel.Range range = sheet.Cells[1, sheet.Columns.Count];
            int lastColumn = range.get_End(Excel.XlDirection.xlToLeft).Column;
            return lastColumn + 1;
        }

        private bool WorksheetExists(Excel.Sheets xlSheets, string sheetName)
        {
            if (xlSheets == null)
            {
                return false;
            }

            object missing = System.Reflection.Missing.Value;

            foreach (Excel.Worksheet sheet in xlSheets)
            {
                if (sheet.Name == sheetName)
                {
                    return true;
                }
            }

            return false;
        }

        // Finds the worksheet with name sheetName and returns it. If it does not exist, returns null.
        private Excel.Worksheet GetWorksheet(Excel.Sheets xlSheets, string sheetName)
        {
            if (xlSheets == null)
            {
                return null;
            }

            object missing = System.Reflection.Missing.Value;

            foreach (Excel.Worksheet sheet in xlSheets)
            {
                if (sheet.Name == sheetName)
                {
                    return sheet;
                }
            }

            return null;
        }

        // Creates and returns a worksheet with the name sheetName.
        private Excel.Worksheet CreateWorksheet(Excel.Sheets xlSheets, string sheetName)
        {
            if (xlSheets == null)
            {
                return null;
            }

            object missing = System.Reflection.Missing.Value;

            int sheetCount = 0;

            foreach (Excel.Worksheet sheet in xlSheets)
            {
                sheetCount++;
            }

            Excel.Worksheet newSheet = xlSheets.Add(xlSheets[sheetCount], missing, missing, missing);
            newSheet.Name = string.Format("{0}", sheetName);

            return newSheet;
        }

        private void RemoveExcelHighlights(Excel.Workbook xlBook)
        {
            if (xlBook == null)
            {
                return;
            }

            foreach (Excel.Worksheet sheet in xlBook.Worksheets)
            {
                Excel.Range range = sheet.UsedRange;
                range.Interior.ColorIndex = 0;
            }
        }

        // Returns true if an IEnumerable<T> is not null and not empty.
        public static bool IsNullOrEmpty<T>(IEnumerable<T> data)
        {
            return data == null || !data.Any();
        }
    }
}
