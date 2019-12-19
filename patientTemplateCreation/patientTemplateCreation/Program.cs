using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;       //microsoft Excel 14 object in references-> COM tab
using System.Collections.Generic;
using System.Windows.Forms;
using System.Diagnostics;

namespace patientTemplateCreation
{
    public class Read_From_Excel
    {
        [STAThread]
        static void Main(string[] args)
        {

            Console.WriteLine("Last Updated: 2019-12-17 by Tennant, Quinton");

            string workingLocExternal = @"\\dc3-pr-files\MedPhysics Backup\Data Extractions\Automation\Info Transfer External.txt";
            string structure = "";
            string planSumExe = "";

            if (new FileInfo(workingLocExternal).Length != 0)
            {
                structure = File.ReadLines(workingLocExternal).Skip(0).Take(1).First();
                //string planSumExe = "";
                planSumExe = File.ReadLines(workingLocExternal).Skip(1).Take(1).First();
            }

			string property = "";
			if (File.ReadAllText(workingLocExternal).Contains("propertyDose"))
			{
				property = "Dose";
			}
			else if (File.ReadAllText(workingLocExternal).Contains("propertyDVH"))
			{
				property = "DVH";
			}
			else
			{
				property = "";
			}

            string inputInfo = $@"\\dc3-pr-files\MedPhysics Backup\Data Extractions\Automation\{property}\Input Information Text Files\{structure} {property} Input Information.txt";

            string planSumRes = "";

            if (planSumExe.Contains("plansumYes"))
            {
                planSumRes = "yes";
            }
            else if (planSumExe.Contains("plansumNo"))
            {
                planSumRes = "no";
            }
            else
            {
                DialogResult result = MessageBox.Show("Would you like to run all patients with Plan Sum included?", "Plan Sum", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    planSumRes = "yes";
                }
                else if (result == DialogResult.No)
                {
                    planSumRes = "no";
                }
            }

            if (!File.Exists(inputInfo))
            {

                //recalls the last location that the user selected for file location.
                inputInfo = patientTemplateCreation.Properties.Settings.Default.textFilePath;
                OpenFileDialog openFileDialog = FormExtensions.CreateOpenFileDialog();

                //conditional if that last location exists then it will start there
                if (!string.IsNullOrWhiteSpace(inputInfo) && File.Exists(inputInfo))
                {
                    openFileDialog.InitialDirectory = inputInfo;
                }
                //if not then it will open the current directory
                else
                {
                    openFileDialog.InitialDirectory = Directory.GetCurrentDirectory();
                }

                while (true)
                {
                    //if user selects a location
                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    { 
                        //Get the path of specified file
                        inputInfo = openFileDialog.FileName;
                        if (Path.GetExtension(inputInfo) == ".txt")
                        {
                            patientTemplateCreation.Properties.Settings.Default.textFilePath = inputInfo;
                            patientTemplateCreation.Properties.Settings.Default.Save();

                            if (inputInfo.Contains("Prostate"))
                            {
                                structure = "Prostate";
                            }
                            else if (inputInfo.Contains("Lung"))
                            {
                                structure = "Lung";
                            }
                            else if (inputInfo.Contains("Breast"))
                            {
                                structure = "Breast";
                            }
                            else if (inputInfo.Contains("Brain"))
                            {
                                structure = "Brain";
                            }
                            else
                            {
                                structure = "";
                            }

                            break;
                        }
                        //if not a text file
                        else
                        {
                            MessageBox.Show("You must select a text file to run the script.", "Invalid File Type",
                                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    //if no location is selected
                    else
                    {
                        MessageBox.Show($"No Input File Selected.", "Failed",
                                                MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }

            //string inputInfo = @"\\dc3-pr-files\MedPhysics Backup\Data Extractions\Automation\Input Information.txt";
            //Get locations for input and output of patient information
            string conFile = File.ReadLines(inputInfo).Skip(4).Take(1).First();
            string excelLoc = File.ReadLines(inputInfo).Skip(8).Take(1).First();
            string patientLoc = File.ReadLines(inputInfo).Skip(24).Take(1).First();
            string outputLoc = File.ReadLines(inputInfo).Skip(16).Take(1).First();

            Directory.CreateDirectory(outputLoc);

            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApplCon = new Excel.Application();
            Excel.Workbook xlWorkbooksCon = xlApplCon.Workbooks.Open(conFile);
            Excel._Worksheet xlWorksheetsCon = xlWorkbooksCon.Sheets[1];
            Excel.Range xlRangeCon = xlWorksheetsCon.UsedRange;

            List<string> listA = new List<string>();
            List<string> listB = new List<string>();
            List<string> listC = new List<string>();

            int rowCountCon = xlRangeCon.Rows.Count;
            int colCountCon = xlRangeCon.Columns.Count;

            //iterate over the rows and columns and print to the console as it appears in the file
            if ((xlRangeCon.Cells[7, 1]) != null || (xlRangeCon.Cells[7, 1]) != null)
            {
                if ((xlRangeCon.Cells[9, 2]) != null || (xlRangeCon.Cells[9, 2]) != null)
                {
                    for (int i = 6; i <= rowCountCon; i++)
                    {
                        listA.Add(xlRangeCon.Cells[i, 1].Value2.ToString());
                        listB.Add(xlRangeCon.Cells[i, 2].Value2.ToString());
                        if (i < 9) { listC.Add(xlRangeCon.Cells[i + 1, 3].Value2.ToString()); }
                    }
                }
                else
                {
                    DialogResult result = MessageBox.Show("Conditional Template not completed.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    if (result == DialogResult.OK)
                    {
                        return;
                    }
                }
            }
            else
            {
                DialogResult result = MessageBox.Show("Conditional Template not completed.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                if (result == DialogResult.OK)
                {
                    return;
                }
            }

            string colArow6 = (listA.ElementAt(0));
            string colArow7 = (listA.ElementAt(1));
            string colArow8 = (listA.ElementAt(2));
            string colArow9 = (listA.ElementAt(3));
            string colBrow6 = (listB.ElementAt(0));
            string colBrow7 = (listB.ElementAt(1));
            string colBrow8 = (listB.ElementAt(2));
            string colBrow9 = (listB.ElementAt(3));
            string colCrow7 = (listC.ElementAt(0));
            string colCrow8 = (listC.ElementAt(1));
            string colCrow9 = (listC.ElementAt(2));

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApplPat = new Excel.Application();
            Excel.Workbook xlWorkbooksPat = xlApplPat.Workbooks.Open(excelLoc);
            Excel._Worksheet xlWorksheetsPat = xlWorkbooksPat.Sheets[1];
            Excel.Range xlRangePat = xlWorksheetsPat.UsedRange;

            //length and width of excel file
            int rowCountPat = xlRangePat.Rows.Count;
            int colCountPat = xlRangePat.Columns.Count;
            //empty the patient ID text file of any previous text
            File.WriteAllText(patientLoc, "");

            //if patient ID's that are being used don't have all 8-characters it will rewrite them to have zeroes to fill to 8.
            for (int i = 2; i <= rowCountPat; i++)
            {
                string cell;
                cell = xlRangePat.Cells[i, 1].Value2.ToString();
                if (cell.Length < 8)
                {
                    Console.Write("\rCreating usable patient IDs                  ");
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
            }
            xlApplPat.DisplayAlerts = false;
            xlWorkbooksPat.Save();

            //iterate over the rows (excel is not zero-based)
            for (int i = 2; i <= rowCountPat; i++)
            {
                //iterate over the columns (excel is not zero-based)
                for (int j = 1; j <= 3; j++)
                {

                    //writes patient IDs into text document
                    if (j == 1)
                    {
                        using (System.IO.StreamWriter PatientList =
                        new System.IO.StreamWriter(patientLoc, true))
                        {
                            PatientList.WriteLine(xlRangePat.Cells[i, j].Value2.ToString());
                        }
                    }
                }
            }

            /* the way the excel file is setup has multiples of each ID
             * so this code deletes any duplicate IDs in the text file
             */
            String[] TextFileLines = File.ReadAllLines(patientLoc);
            String[] TextFileLinesDist;
            TextFileLinesDist = TextFileLines.Distinct().ToArray();
            File.WriteAllLines(patientLoc, TextFileLinesDist);

            string usableID;

            // Read the file line by line.  
            System.IO.StreamReader file =
                new System.IO.StreamReader(patientLoc);
            //x is the count of which row we are on per loop so it starts at 2 to skip header
            int x = 2;
            while ((usableID = file.ReadLine()) != null && x<=rowCountPat)
            {
                //uses the ID of the current patient to write a template text file into Templates folder
                string templateIDLoc = File.ReadLines(inputInfo).Skip(16).Take(1).First();

                Directory.CreateDirectory(templateIDLoc);

                string templateLoc = $@"{templateIDLoc}\{usableID}.txt";
                File.WriteAllText(templateLoc, String.Empty);

                //Template printer
                using (System.IO.StreamWriter patientTemplate =
                new System.IO.StreamWriter(templateLoc, true))
                {
                    //Get's file location from Blank template
                    string blankLoc = File.ReadLines(inputInfo).Skip(12).Take(1).First();
                    patientTemplate.WriteLine(File.ReadLines(inputInfo).Skip(40).Take(1).First());
                    patientTemplate.WriteLine(""); //spacing

                    int lineCount = File.ReadLines(blankLoc).Count();

                    /* iterate over ID rows
                     * but issue with x<rowCount that it doesn't include last row for some reason
                     * to fix must have x<=rowCount but in that case it says cant operate on null??
                     */
					 //DOSE DOSE DOSE
                    if ((xlRangePat.Cells[x, 1].Value2 != null || xlRangePat.Cells[x, 1].Value2 != "") && property == "Dose" && xlRangePat.Cells[x, 1].Value2 is string)
                    {
						while (xlRangePat.Cells[x, 1].Value2 == usableID && x <= rowCountPat)
						{
							patientTemplate.WriteLine(xlRangePat.Cells[x, 2].Value2.ToString());
							//conditional statements for identifying what Nurses have written in the system (not exhaustive so may need to expand)
							//for when given includes both PROS and PELB we want to print PELB
							if ((xlRangePat.Cells[x, 3].Value2.ToString()).Contains(colArow6) == (colArow7.Contains("T")) && (xlRangePat.Cells[x, 3].Value2.ToString()).Contains(colBrow6) == (colBrow7.Contains("T")))
							{
								patientTemplate.WriteLine($"{colCrow7} | " + xlRangePat.Cells[x, 3].Value2.ToString());
							}
							//for when given includes PROS but not PELB we want to print PROS
							else if ((xlRangePat.Cells[x, 3].Value2.ToString()).Contains(colArow6) == (colArow8.Contains("T")) && (xlRangePat.Cells[x, 3].Value2.ToString()).Contains(colBrow6) == (colBrow8.Contains("T")))
							{
								patientTemplate.WriteLine($"{colCrow8} | " + xlRangePat.Cells[x, 3].Value2.ToString());
							}
							//for when given includes PELB but not PROS we want to print PELB
							else if ((xlRangePat.Cells[x, 3].Value2.ToString()).Contains(colArow6) == (colArow9.Contains("T")) && (xlRangePat.Cells[x, 3].Value2.ToString()).Contains(colBrow6) == (colBrow9.Contains("T")))
							{
								patientTemplate.WriteLine($"{colCrow9} | " + xlRangePat.Cells[x, 3].Value2.ToString());
							}

							//iterating over blank template to copy the structures within the Plan, but needs to be expanded for individual cases
							int count = 0;
							if (xlRangePat.Cells[x, 3].Value2.ToString().Contains("LUNR"))
							{
								count = count + 4;
							}
							else if (xlRangePat.Cells[x, 3].Value2.ToString().Contains("LUNL"))
							{
								count = count + 18;
							}
							else if (xlRangePat.Cells[x, 3].Value2.ToString().Contains("BRER") || xlRangePat.Cells[x, 3].Value2.ToString().Contains("SCNR"))
							{
								count = count + 21;
							}
							else if (xlRangePat.Cells[x, 3].Value2.ToString().Contains("BREL") || xlRangePat.Cells[x, 3].Value2.ToString().Contains("SCNL"))
							{
								count = count + 4;
							}
							else if (xlRangePat.Cells[x, 3].Value2.ToString().Contains("BRAI") && xlRangePat.Cells[x, 3].Value2.ToString().Contains("1"))
							{
								count = count + 22;
							}
							else if (xlRangePat.Cells[x, 3].Value2.ToString().Contains("BRAI") && xlRangePat.Cells[x, 3].Value2.ToString().Contains("2"))
							{
								count = count + 40;
							}
							else if (xlRangePat.Cells[x, 3].Value2.ToString().Contains("BRAI") && xlRangePat.Cells[x, 3].Value2.ToString().Contains("3"))
							{
								count = count + 58;
							}
							else if (xlRangePat.Cells[x, 3].Value2.ToString().Contains("BRAI") && xlRangePat.Cells[x, 3].Value2.ToString().Contains("4"))
							{
								count = count + 76;
							}
							else if (xlRangePat.Cells[x, 3].Value2.ToString().Contains("BRAI") && xlRangePat.Cells[x, 3].Value2.ToString().Contains("5"))
							{
								count = count + 94;
							}
							else if (xlRangePat.Cells[x, 3].Value2.ToString().Contains("BRAI") && xlRangePat.Cells[x, 3].Value2.ToString().Contains("6"))
							{
								count = count + 112;
							}
							else if (xlRangePat.Cells[x, 3].Value2.ToString().Contains("BRAI") && xlRangePat.Cells[x, 3].Value2.ToString().Contains("7"))
							{
								count = count + 130;
							}
							else
							{
								count = count + 4;
							}
							for (; count < lineCount; count++)
							{
								if (File.ReadLines(blankLoc).Skip(count).Take(1).First() == null || File.ReadLines(blankLoc).Skip(count).Take(1).First() == "")
								{
									break;
								}
								patientTemplate.WriteLine(File.ReadLines(blankLoc).Skip(count).Take(1).First());
							}

							patientTemplate.WriteLine(""); //spacing
                            //creating fraction of completed template percentage since I prefer to see if the code is actually running what I want it to
                            float perLoad = ((float)x / (float)rowCountPat) * 100.00F;
                            double perLoadR = Math.Round(perLoad, 0);
                            string perLoadString = perLoadR.ToString();
                            Console.Write($"\r Creating {structure} {property} Templates: {perLoadString}%     ");
                            //increment x
                            x++;
                        }
                    }
					//DVH DVH DVH
					else if ((xlRangePat.Cells[x, 1].Value2 != null || xlRangePat.Cells[x, 1].Value2 != "") && property == "DVH")
					{
						while (xlRangePat.Cells[x, 1].Value2 == usableID && x <= rowCountPat)
						{
							patientTemplate.WriteLine(xlRangePat.Cells[x, 2].Value2.ToString());
							patientTemplate.WriteLine(xlRangePat.Cells[x, 3].Value2.ToString());
							////conditional statements for identifying what Nurses have written in the system (not exhaustive so may need to expand)
							////for when given includes both PROS and PELB we want to print PELB
							//if ((xlRangePat.Cells[x, 3].Value2.ToString()).Contains(colArow6) == (colArow7.Contains("T")) && (xlRangePat.Cells[x, 3].Value2.ToString()).Contains(colBrow6) == (colBrow7.Contains("T")))
							//{
							//	patientTemplate.WriteLine($"{colCrow7}");
							//}
							////for when given includes PROS but not PELB we want to print PROS
							//else if ((xlRangePat.Cells[x, 3].Value2.ToString()).Contains(colArow6) == (colArow8.Contains("T")) && (xlRangePat.Cells[x, 3].Value2.ToString()).Contains(colBrow6) == (colBrow8.Contains("T")))
							//{
							//	patientTemplate.WriteLine($"{colCrow8}");
							//}
							////for when given includes PELB but not PROS we want to print PELB
							//else if ((xlRangePat.Cells[x, 3].Value2.ToString()).Contains(colArow6) == (colArow9.Contains("T")) && (xlRangePat.Cells[x, 3].Value2.ToString()).Contains(colBrow6) == (colBrow9.Contains("T")))
							//{
							//	patientTemplate.WriteLine($"{colCrow9}");
							//}

							//iterating over blank template to copy the structures within the Plan, but needs to be expanded for individual cases
							int count = 0;
							if (xlRangePat.Cells[x, 3].Value2.ToString().Contains("LUNR"))
							{
								count = count + 18;

							}
							else if (xlRangePat.Cells[x, 3].Value2.ToString().Contains("LUNL"))
							{
								count = count + 4;

							}
							//else if (xlRangePat.Cells[x, 3].Value2.ToString().Contains("BRER") || xlRangePat.Cells[x, 3].Value2.ToString().Contains("SCNR"))
							//{
							//	count = count + 10;
							//}
							//else if (xlRangePat.Cells[x, 3].Value2.ToString().Contains("BREL") || xlRangePat.Cells[x, 3].Value2.ToString().Contains("SCNL"))
							//{
							//	count = count + 4;
							//}
							else if (xlRangePat.Cells[x, 3].Value2.ToString().Contains("BRAI") && xlRangePat.Cells[x, 3].Value2.ToString().Contains("1"))
							{
								count = count + 18;
							}
							else if (xlRangePat.Cells[x, 3].Value2.ToString().Contains("BRAI") && xlRangePat.Cells[x, 3].Value2.ToString().Contains("2"))
							{
								count = count + 32;
							}
							else if (xlRangePat.Cells[x, 3].Value2.ToString().Contains("BRAI") && xlRangePat.Cells[x, 3].Value2.ToString().Contains("3"))
							{
								count = count + 46;
							}
							else if (xlRangePat.Cells[x, 3].Value2.ToString().Contains("BRAI") && xlRangePat.Cells[x, 3].Value2.ToString().Contains("4"))
							{
								count = count + 60;
							}
							else if (xlRangePat.Cells[x, 3].Value2.ToString().Contains("BRAI") && xlRangePat.Cells[x, 3].Value2.ToString().Contains("5"))
							{
								count = count + 74;
							}
							else if (xlRangePat.Cells[x, 3].Value2.ToString().Contains("BRAI") && xlRangePat.Cells[x, 3].Value2.ToString().Contains("6"))
							{
								count = count + 88;
							}
							else if (xlRangePat.Cells[x, 3].Value2.ToString().Contains("BRAI") && xlRangePat.Cells[x, 3].Value2.ToString().Contains("7"))
							{
								count = count + 102;
							}
							else
							{
								count = count + 4;
							}

							for (; count < lineCount; count++)
							{
								if (File.ReadLines(blankLoc).Skip(count).Take(1).First() == null || File.ReadLines(blankLoc).Skip(count).Take(1).First() == "")
								{
									break;
								}
								patientTemplate.WriteLine(File.ReadLines(blankLoc).Skip(count).Take(1).First());
							}

							patientTemplate.WriteLine(""); //spacing
														   //creating fraction of completed template percentage since I prefer to see if the code is actually running what I want it to
							float perLoad = ((float)x / (float)rowCountPat) * 100.00F;
							double perLoadR = Math.Round(perLoad, 0);
							string perLoadString = perLoadR.ToString();
							Console.Write($"\r Creating {structure} {property} Templates: {perLoadString}%     ");
							//increment x
							x++;
						}
					}

					if (planSumRes == "yes" && property == "DVH")
                    {
                        for (int track = 30; track <= 42; track++)
                        {
                            patientTemplate.WriteLine(File.ReadLines(blankLoc).Skip(track).Take(1).First());
                        }
                    }
					else if (planSumRes == "yes" && property == "Dose")
					{
						for (int track = 26; track <= 36; track++)
						{
							patientTemplate.WriteLine(File.ReadLines(blankLoc).Skip(track).Take(1).First());
						}
					}
                }
            }
            //close the file
            file.Close();

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(xlRangePat);
            Marshal.ReleaseComObject(xlWorksheetsPat);

            Marshal.ReleaseComObject(xlRangeCon);
            Marshal.ReleaseComObject(xlWorksheetsCon);

            //close and release
            xlWorkbooksPat.Close();
            Marshal.ReleaseComObject(xlWorkbooksPat);

            xlWorkbooksCon.Close();
            Marshal.ReleaseComObject(xlWorkbooksCon);

            //quit and release
            xlApplPat.Quit();
            Marshal.ReleaseComObject(xlApplPat);

            xlApplCon.Quit();
            Marshal.ReleaseComObject(xlApplCon);

            Console.Write("\n\r Copying");

            //code to copy patient IDs from in the Templates folder to into where it is taken for running TestDoseExtraction
            string from = File.ReadLines(inputInfo).Skip(24).Take(1).First(); ;
            string to = File.ReadLines(inputInfo).Skip(32).Take(1).First();
            string idListDir = to.Substring(0, 71);
            File.Copy(from, to, true);
            //string idList = $@"{idListDir}\{structure} Pending Extractions List.txt";
            //File.Copy(to, idList);

            Console.Write("\r Copying.   ");

            //get todays date and time
            DateTime localDate = DateTime.Now;
            string localDateString = localDate.ToString("yyyy-MM-dd   hh mm tt");

            Console.Write("\r Copying..    ");

            //create archive of templates any time it's run just in case
            string sourceDir = File.ReadLines(inputInfo).Skip(16).Take(1).First();
            string targetDir = File.ReadLines(inputInfo).Skip(20).Take(1).First();
            string timeTargetDir = $@"{targetDir}\{structure} {property} - {localDateString}";

            if (!Directory.Exists(timeTargetDir))
            {
                Directory.CreateDirectory(timeTargetDir);
            }
            foreach (var srcPath in Directory.GetFiles(sourceDir))
            {
                //Copy the file from sourcepath and place into mentioned target path, 
                //Overwrite the file if same file is exist in target path
                File.Copy(srcPath, srcPath.Replace(sourceDir, timeTargetDir), true);
            }

            Console.Write("\r Copying...     \n");

            //string sourceFile = @"\\dc3-pr-files\MedPhysics Backup\Coop Students\2019\Term 3 - Autumn\Quinton Tennant\Release Copy\Clear Templates Archive.lnk";
            //string targetFile = $@"{targetDir}\Clear Template Archives.lnk";
			//
            //if (!File.Exists(targetFile))
            //{
            //    File.Copy(sourceFile, targetFile);
            //}
        }
    }
}