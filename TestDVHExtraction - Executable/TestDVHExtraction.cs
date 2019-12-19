////////////////////////////////////////////////////////////////////////////////
// TestDVHExtraction.cs
//
//  A ESAPI v11+ script that demonstrates DVH export.
//
// Copyright (c) 2015 Varian Medical Systems, Inc.
//
// Permission is hereby granted, free of charge, to any person obtaining a copy 
// of this software and associated documentation files (the "Software"), to deal 
// in the Software without restriction, including without limitation the rights 
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell 
// copies of the Software, and to permit persons to whom the Software is 
// furnished to do so, subject to the following conditions:
//
//  The above copyright notice and this permission notice shall be included in 
//  all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR 
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, 
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL 
// THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER 
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, 
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN 
// THE SOFTWARE.
////////////////////////////////////////////////////////////////////////////////
using System;
using System.Linq;
using System.Text;
using System.Collections.Generic;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;
using System.IO;
using System.Windows.Forms;
using Application = VMS.TPS.Common.Model.API.Application;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace TestDVHExtraction
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
			uint timeout = 0;
            try
            {
                using (Application app = Application.CreateApplication("studentx", "Studentx01"))
                {
                    Execute(app);
                }
            }
            catch (Exception e)
            {
				MessageBoxEx.Show("Unhandled exception: " + e.ToString(), "Error",
					MessageBoxButtons.OK, MessageBoxIcon.Error, timeout);
				Console.Error.WriteLine(e.ToString());
            }
        }

        static void Execute(Application app)
        {
			string workingLocInternal = @"\\dc3-pr-files\MedPhysics Backup\Data Extractions\Automation\Info Transfer Internal.txt";
			string workingLocExternal = @"\\dc3-pr-files\MedPhysics Backup\Data Extractions\Automation\Info Transfer External.txt";


			Console.WriteLine("Last Updated: 2019-12-17 by Tennant, Quinton");

			//Get todays date to identify when the failed extraction document was made
			DateTime localDate = DateTime.Now;
			string localDateString = localDate.ToString("yyyy/MM/dd hh:mm:ss tt");

			DateTime localDateA = DateTime.Now;
			string localTimeStringA = localDateA.ToString("hh:mm:ss tt");
			Console.WriteLine("Started: " + localTimeStringA);

			Values myValues = new Values();
			myValues.getStructure();
			string structure = myValues.structure;

			myValues.getInputInfo();
			string inputInfo = myValues.inputInfo;


			if (!File.Exists(inputInfo))
			{
				//get user input for information file locations.
				inputInfo = Properties.Settings.Default.textFilePath;
			
				OpenFileDialog openFileDialog = FormExtensions.CreateOpenFileDialog();
			
				if (!string.IsNullOrWhiteSpace(inputInfo) && File.Exists(inputInfo))
				{
					openFileDialog.InitialDirectory = inputInfo;
				}
				else
				{
					openFileDialog.InitialDirectory = Directory.GetCurrentDirectory();
				}
			
				while (true)
				{
					if (openFileDialog.ShowDialog() == DialogResult.OK)
					{
						//Get the path of specified file
						inputInfo = openFileDialog.FileName;
						if (Path.GetExtension(inputInfo) == ".txt")
						{
							if (inputInfo.ToLower().Contains("brain"))
							{
								structure = "Brain";
							}
							else if (inputInfo.ToLower().Contains("breast"))
							{
								structure = "Breast";
							}
							else if (inputInfo.ToLower().Contains("lung"))
							{
								structure = "Lung";
							}
							else if (inputInfo.ToLower().Contains("prostate"))
							{
								structure = "Prostate";
							}
							else
							{
								structure = "Custom";
							}

							Properties.Settings.Default.textFilePath = inputInfo;
							Properties.Settings.Default.Save();
							break;
						}
						else
						{
							MessageBox.Show("You must select a text file to run the script.", "Invalid File Type",
											MessageBoxButtons.OK, MessageBoxIcon.Error);
						}
					}
					else
					{
						MessageBox.Show($"No Input File Selected.", "Failed",
												MessageBoxButtons.OK, MessageBoxIcon.Error);
						return;
					}
				}
			}

			string singleID = "";
			if (File.ReadLines(workingLocExternal).Count() == 3)
			{
				if (File.ReadLines(workingLocExternal).Skip(2).Take(1).First().Length == 8)
				{
					singleID = File.ReadLines(workingLocExternal).Skip(2).Take(1).First();
				}
			}

			//Getting input and output information from input information text file so users don't have to enter coding env
			//Working file is used as a roundabout way to pass patient ID DVHQuery.cs (if you can find a better way, go ahead!)
			string filePath = File.ReadLines(inputInfo).Skip(40).Take(1).First();
			string failedExtractionLoc = File.ReadLines(inputInfo).Skip(36).Take(1).First();
			string usableIDLoc = File.ReadLines(inputInfo).Skip(32).Take(1).First();
			string textFileLoc = File.ReadLines(inputInfo).Skip(16).Take(1).First();
			string usableID;

			uint timeout = 0; //messagebox timeout timer
			int count = 0;
			int saveCount = 0;

			//Empties working file and writes the created date header in failed extractions
			File.WriteAllText(@failedExtractionLoc, "Created " + localDateString + System.Environment.NewLine);


			Microsoft.Office.Interop.Excel.Application oXL;
			Microsoft.Office.Interop.Excel._Workbook oWB;
			Microsoft.Office.Interop.Excel._Worksheet oSheet;

			object misvalue = System.Reflection.Missing.Value;
			//Start Excel and get Application object.
			oXL = new Microsoft.Office.Interop.Excel.Application();
			oXL.DisplayAlerts = false;
			oXL.Visible = true;

			//Get a new workbook.
			if (File.Exists(filePath))
			{
				oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Open(filePath));
			}
			else
			{
				oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Add(""));
			}
			oWB.SaveAs(filePath, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
				false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
				Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
			//oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;

			string sampleTemplate = File.ReadLines(inputInfo).Skip(12).Take(1).First();
			int lineCount = File.ReadLines(sampleTemplate).Count();
			List<string> subStruc = new List<string>();
			List<string> sheetNames = new List<string>();

			foreach (Excel.Worksheet sheet in oWB.Sheets)
			{
				sheetNames.Add(sheet.Name);
			}

			var xlSheets = oWB.Sheets as Excel.Sheets;
			int sheetCount = 1;
			for (int sampleCount = 5; sampleCount < lineCount; sampleCount++)
			{
				if (File.ReadLines(sampleTemplate).Skip(sampleCount).Take(1).First() == null || File.ReadLines(sampleTemplate).Skip(sampleCount).Take(1).First() == "")
				{
					break;
				}
				subStruc.Add(File.ReadLines(sampleTemplate).Skip(sampleCount).Take(1).First());

				if (File.ReadAllLines(sampleTemplate).ToString().ToUpper().Contains("PLAN SUM"))
				{
					subStruc.Add("Plan Sum");
					if (!sheetNames.Contains("Plan Sum"))
					{
						var xlSheetPlanSum = (Excel.Worksheet)xlSheets.Add(xlSheets[sheetCount], Type.Missing, Type.Missing, Type.Missing);
						xlSheetPlanSum.Name = "Plan Sum";
						sheetCount++;
					}
				}

				if (!sheetNames.Contains($"{subStruc[sampleCount - 5]}"))
				{
					var xlSheetNew = (Excel.Worksheet)xlSheets.Add(xlSheets[sheetCount], Type.Missing, Type.Missing, Type.Missing);
					xlSheetNew.Name = (File.ReadLines(sampleTemplate).Skip(sampleCount).Take(1).First());
					xlSheetNew.Cells[2, 1] = "mean";
					xlSheetNew.Cells[2, 2] = "min";
					xlSheetNew.Cells[2, 3] = "max";
					xlSheetNew.Cells[2, 4] = "stdev";
					xlSheetNew.Cells[1, 6] = "Dose [cGy]";
					xlSheetNew.Cells[1, 6].Font.Bold = true;
					sheetCount++;
					if (xlSheetNew.Name.Contains("PTV_LUNL"))
					{
						subStruc.Add("PTV_LUNR");
						xlSheetNew = (Excel.Worksheet)xlSheets.Add(xlSheets[sheetCount], Type.Missing, Type.Missing, Type.Missing);
						xlSheetNew.Name = ("PTV_LUNR");
						xlSheetNew.Cells[2, 1] = "mean";
						xlSheetNew.Cells[2, 2] = "min";
						xlSheetNew.Cells[2, 3] = "max";
						xlSheetNew.Cells[2, 4] = "stdev";
						xlSheetNew.Cells[1, 6] = "Dose [cGy]";
						xlSheetNew.Cells[1, 6].Font.Bold = true;
						sheetCount++;
					}

					if (xlSheetNew.Name.Contains("GTV_BRAI") && !xlSheetNew.Name.Contains("TOTAL"))
                    {
						subStruc.Add("GTV_BRAI1");
						var xlSheetSubNew = (Excel.Worksheet)xlSheets.Add(xlSheets[sheetCount], Type.Missing, Type.Missing, Type.Missing);
                        xlSheetSubNew.Name = ("GTV_BRAI1");
                        xlSheetSubNew.Cells[2, 1] = "mean";
                        xlSheetSubNew.Cells[2, 2] = "min";
                        xlSheetSubNew.Cells[2, 3] = "max";
                        xlSheetSubNew.Cells[2, 4] = "stdev";
                        xlSheetSubNew.Cells[1, 6] = "Dose [cGy]";
                        xlSheetSubNew.Cells[1, 6].Font.Bold = true;
						sheetCount++;

						subStruc.Add("GTV_BRAI2");
						xlSheetSubNew = (Excel.Worksheet)xlSheets.Add(xlSheets[sheetCount], Type.Missing, Type.Missing, Type.Missing);
                        xlSheetSubNew.Name = ("GTV_BRAI2");
                        xlSheetSubNew.Cells[2, 1] = "mean";
                        xlSheetSubNew.Cells[2, 2] = "min";
                        xlSheetSubNew.Cells[2, 3] = "max";
						xlSheetSubNew.Cells[2, 4] = "stdev";
						xlSheetSubNew.Cells[1, 6] = "Dose [cGy]";
						xlSheetSubNew.Cells[1, 6].Font.Bold = true;
						sheetCount++;

						subStruc.Add("GTV_BRAI3");
						xlSheetSubNew = (Excel.Worksheet)xlSheets.Add(xlSheets[sheetCount], Type.Missing, Type.Missing, Type.Missing);
						xlSheetSubNew.Name = ("GTV_BRAI3");
						xlSheetSubNew.Cells[2, 1] = "mean";
						xlSheetSubNew.Cells[2, 2] = "min";
						xlSheetSubNew.Cells[2, 3] = "max";
						xlSheetSubNew.Cells[2, 4] = "stdev";
						xlSheetSubNew.Cells[1, 6] = "Dose [cGy]";
						xlSheetSubNew.Cells[1, 6].Font.Bold = true;
						sheetCount++;

						subStruc.Add("GTV_BRAI4");
						xlSheetSubNew = (Excel.Worksheet)xlSheets.Add(xlSheets[sheetCount], Type.Missing, Type.Missing, Type.Missing);
						xlSheetSubNew.Name = ("GTV_BRAI4");
						xlSheetSubNew.Cells[2, 1] = "mean";
						xlSheetSubNew.Cells[2, 2] = "min";
						xlSheetSubNew.Cells[2, 3] = "max";
						xlSheetSubNew.Cells[2, 4] = "stdev";
						xlSheetSubNew.Cells[1, 6] = "Dose [cGy]";
						xlSheetSubNew.Cells[1, 6].Font.Bold = true;
						sheetCount++;

						subStruc.Add("GTV_BRAI5");
						xlSheetSubNew = (Excel.Worksheet)xlSheets.Add(xlSheets[sheetCount], Type.Missing, Type.Missing, Type.Missing);
						xlSheetSubNew.Name = ("GTV_BRAI5");
						xlSheetSubNew.Cells[2, 1] = "mean";
						xlSheetSubNew.Cells[2, 2] = "min";
						xlSheetSubNew.Cells[2, 3] = "max";
						xlSheetSubNew.Cells[2, 4] = "stdev";
						xlSheetSubNew.Cells[1, 6] = "Dose [cGy]";
						xlSheetSubNew.Cells[1, 6].Font.Bold = true;
						sheetCount++;

						subStruc.Add("GTV_BRAI6");
						xlSheetSubNew = (Excel.Worksheet)xlSheets.Add(xlSheets[sheetCount], Type.Missing, Type.Missing, Type.Missing);
						xlSheetSubNew.Name = ("GTV_BRAI6");
						xlSheetSubNew.Cells[2, 1] = "mean";
						xlSheetSubNew.Cells[2, 2] = "min";
						xlSheetSubNew.Cells[2, 3] = "max";
						xlSheetSubNew.Cells[2, 4] = "stdev";
						xlSheetSubNew.Cells[1, 6] = "Dose [cGy]";
						xlSheetSubNew.Cells[1, 6].Font.Bold = true;
						sheetCount++;

						subStruc.Add("GTV_BRAI7");
						xlSheetSubNew = (Excel.Worksheet)xlSheets.Add(xlSheets[sheetCount], Type.Missing, Type.Missing, Type.Missing);
						xlSheetSubNew.Name = ("GTV_BRAI7");
						xlSheetSubNew.Cells[2, 1] = "mean";
						xlSheetSubNew.Cells[2, 2] = "min";
						xlSheetSubNew.Cells[2, 3] = "max";
						xlSheetSubNew.Cells[2, 4] = "stdev";
						xlSheetSubNew.Cells[1, 6] = "Dose [cGy]";
						xlSheetSubNew.Cells[1, 6].Font.Bold = true;
						sheetCount++;
					}
				}
			}


			if (File.Exists(usableIDLoc))
			{
				List<string> planIds = new List<string>();
				string extractedDVH = $@"\\dc3-pr-files\MedPhysics Backup\Data Extractions\Extracted DVH\{structure}";

				//get patient ID list from the ID location given in input information
				System.IO.StreamReader file =
					new System.IO.StreamReader(@usableIDLoc);
				//iterate entire code over each ID
				while ((usableID = file.ReadLine()) != null && count == 0)
				{
					if (singleID != "")
					{
						count++;
						if (count > 0)
						{
							usableID = singleID;
						}
					}

					//writes over all text in the working.txt to transfer ID to DVHQuery.cs
					File.WriteAllText(workingLocInternal, "");
					File.WriteAllText(workingLocInternal, inputInfo + Environment.NewLine);
					File.AppendAllText(workingLocInternal, usableID + Environment.NewLine);
					//Make sure any patients with studentx are closed otherwise code won't run
					app.ClosePatient();
					// Load the patient.
					Patient patient = app.OpenPatientById(usableID);
					if (patient == null)
					{
					    MessageBoxEx.Show("No patient is loaded.", "Extraction Failed",
							MessageBoxButtons.OK, MessageBoxIcon.Error, timeout);

						//When extraction fails write the patient ID and reason to failed extractions
						using (System.IO.StreamWriter failedExtractionFile =
							new System.IO.StreamWriter(failedExtractionLoc, true))
						{
							failedExtractionFile.Write(usableID);
							failedExtractionFile.WriteLine(" -- No viable patient was loaded.");
						}
					}

					string textFilePath = $@"{textFileLoc}\{usableID}.txt";
					string saveDirectory = "";
					List<Tuple<string, string, string, DVHData>> dvhList = new List<Tuple<string, string, string, DVHData>>();


					if (File.Exists(textFilePath))
					{
						Console.Write($"\rRunning DVH Extraction on {usableID}                                  ");
						// Reads each line of the file into a string array, ignoring empty lines.
						string[] lines = ReadNonBlankLines(textFilePath);

						int index = 1; // lines[index] == Course
						int numOfLines = lines.Length;

						planIds.Clear();
						for (; index < numOfLines;)
						{
							// Check if the Patient has Courses. If not, there is no DVH data to extract
							// so we return.
							if (IsNullOrEmpty(patient.Courses))
							{
								return;
							}
							Course course = patient.Courses.FirstOrDefault(x => CheckIdMatch(x.Id, lines[index]));
							index++; // lines[index] == PlanSetup or PlanSum
							if (course != null)
							{
								course.Id.ToString();
							}
							else
							{
								//Self-exiting messagebox pop up (change timeout timer at top to change exit time)
								MessageBoxEx.Show($"Course {course} does not exist for {usableID}.", "Extraction Failed",
												MessageBoxButtons.OK, MessageBoxIcon.Error, timeout);
								//if extraction fails writes to failed extractions w/ ID and reason
								using (System.IO.StreamWriter failedExtractionFile =
									new System.IO.StreamWriter(failedExtractionLoc, true))
								{
									failedExtractionFile.Write(usableID);
									failedExtractionFile.WriteLine($" -- Course {course} did not exist.");
								}
							}
							if (course != null)
							{
								// Check if the Course has any PlanSetups/PlanSums. If not, a different Course
								// could have plans, so we continue looping until we get to the next Course in
								// the text file.
								if (IsNullOrEmpty(course.PlanSetups) && IsNullOrEmpty(course.PlanSums))
								{
									// Increments the index so that on the next loop we are immediately at the next Course.
									// (Skips looking at the plans and structures).
									index = index + StringToInt(lines[index + 1]) + 2;
									continue;
								}
								PlanSetup planSetup = course.PlanSetups.FirstOrDefault(x => CheckIdMatch(x.Id, lines[index]));
								PlanSum planSum = course.PlanSums.FirstOrDefault(x => CheckIdMatch(x.Id, lines[index]));
								string planId = lines[index];
								planIds.Add(planId);

								index++; // lines[index] == Number of Queried Structures

								if (planSetup != null)
								{
									// Check if the PlanSetup has a StructureSet/Structures. If not, a
									// different plan could have them, so we continue looping until we get
									// to the next Course in the text file (could be the same Course).
									if (planSetup.StructureSet == null || IsNullOrEmpty(planSetup.StructureSet.Structures))
									{
										// Increments the index so that on the next loop we are immediately at the next Course.
										// (Skips looking at the structures).
										index = index + StringToInt(lines[index]) + 1;
										continue;
									}
									int numOfStructs = StringToInt(lines[index]);
									index++; // lines[index] == Structure

									for (int structCount = 0; structCount < numOfStructs; structCount++)
									{
										Structure target = planSetup.StructureSet.Structures.FirstOrDefault(x => CheckIdMatch(x.Id, lines[index]));

										index++; // lines[index] == Structure or Course or PlanSetup/PlanSum

										if (target != null)
										{
											DVHData dvh = planSetup.GetDVHCumulativeData(target, DoseValuePresentation.Absolute, VolumePresentation.Relative, 10);

											if (dvh != null)
											{
												dvhList.Add(Tuple.Create(course.Id, planSetup.Id, target.Id, dvh));
											}
										}
									}
								}
								else if (planSum != null)
								{
									// Check if the PlanSum has a StructureSet/Structures. If not, a
									// different plan could have them, so we continue looping until we get
									// to the next Course in the text file (could be the same Course).
									if (planSum.StructureSet == null || IsNullOrEmpty(planSum.StructureSet.Structures))
									{
										// Increments the index so that on the next loop we are immediately at the next Course.
										// (Skips looking at the structures).
										index = index + StringToInt(lines[index]) + 1;
										continue;
									}
									int numOfStructs = StringToInt(lines[index]);
									index++; // lines[index] == Structure

									for (int structCount = 0; structCount < numOfStructs; structCount++)
									{
										Structure target = planSum.StructureSet.Structures.FirstOrDefault(x => CheckIdMatch(x.Id, lines[index]));
										index++; // lines[index] == Structure or Course or PlanSetup/PlanSum

										if (target != null)
										{
											DVHData dvh = planSum.GetDVHCumulativeData(target, DoseValuePresentation.Absolute, VolumePresentation.Relative, 10);
											

											if (dvh != null)
											{
												dvhList.Add(Tuple.Create(course.Id, planSum.Id, target.Id, dvh));
											}
										}
									}
								}
							}
						}

						string studyId = GenerateStudyId(patient);

						var activeSheet = oWB.Sheets["Sheet1"];
						int TestCount = 0;
						int planIdIndex = -1;
						string firstStruc = "";

						foreach (Tuple<string, string, string, DVHData> t in dvhList)
						{
							if (TestCount == 0)
							{
								firstStruc = t.Item3;
							}

							if (t.Item3 == firstStruc)
							{
								planIdIndex++;
							}

							string activePlanId = planIds[planIdIndex];
							activeSheet = oWB.Sheets[$"{t.Item3}"];
							activeSheet.Select();

							Excel.Range visibleRange = oXL.ActiveWindow.VisibleRange;
							Excel.Range last = activeSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);

							int colTally = last.Column;
							int colCount = colTally + 1;


							if (colCount < visibleRange.Column || colCount > (visibleRange.Column + visibleRange.Columns.Count - 2))
							{
								// Scroll Excel worksheet if the row we are writing to is offscreen.
								int colToScrollTo = colCount - visibleRange.Columns.Count;

								if (colToScrollTo < 1)
								{
									oXL.ActiveWindow.ScrollColumn = 1;
								}
								else
								{
									oXL.ActiveWindow.ScrollColumn = colToScrollTo + 6;
								}
							}

							//specFilePath = $@"{filePath}\{t.Item2}\Patient DVH {structure} Spreadsheet.xlsx";

							int numOfPoints = t.Item4.CurveData.Length;

							string result;
							string utterance1 = $"{studyId}";
							string utterance2 = $"{activePlanId}";


							//MessageBox.Show(planIdIndex.ToString());

							Excel.Range range1 = activeSheet.Rows["2:2"];
							Excel.Range findRange1;
							Excel.Range findRange1Next1;
							Excel.Range findRange1Next2;
							Excel.Range findRange1Next3;
							Excel.Range findRange1Next4;
							findRange1 = range1.Find(utterance1);

							//result = (activeSheet.Cells[1, findRange.Column] as Excel.Range).Value2.ToString();

							Excel.Range range2 = activeSheet.Rows["1:1"];
							Excel.Range findRange2;
							findRange2 = range2.Find(utterance2);

							if (findRange1 == null)
							{
								activeSheet.Cells[2, colCount].NumberFormat = "@";
								activeSheet.Cells[2, colCount].Font.Bold = true;

								activeSheet.Cells[1, colCount] = activePlanId;
								activeSheet.Cells[2, colCount] = studyId;

								for (int i = 3; i < numOfPoints; i++)
								{
									var lastUsedCol = activeSheet.Range[activeSheet.Cells[1, 6], activeSheet.Cells[last.Row, colCount - 1].EntireColumn];
									lastUsedCol.Interior.ColorIndex = 0;

									if (i > last.Row)
									{
										activeSheet.Cells[i, 6] = t.Item4.CurveData[i - 3].DoseValue.Dose;
									}

									activeSheet.Cells[i, colCount] = t.Item4.CurveData[i - 3].Volume;
									activeSheet.Cells[i, colCount].EntireColumn.Interior.Color = System.Drawing.Color.Yellow;

								}
							}
							else
							{
								string firstRow = activeSheet.Cells[1, findRange1.Column].ToString();
								if (firstRow == utterance2)
								{
									if (findRange1.Column < visibleRange.Column || findRange1.Column > (visibleRange.Column + visibleRange.Columns.Count - 2))
									{
										// Scroll Excel worksheet if the row we are writing to is offscreen.
										int colToScrollTo = findRange1.Column - visibleRange.Columns.Count;

										if (colToScrollTo < 1)
										{
											oXL.ActiveWindow.ScrollColumn = 1;
										}
										else
										{
											oXL.ActiveWindow.ScrollColumn = colToScrollTo + 6;
										}
									}

									var leftUsedCol = activeSheet.Range[activeSheet.Cells[1, 6], activeSheet.Cells[last.Row, findRange1.Column - 1].EntireColumn];
									var rightUsedCol = activeSheet.Range[activeSheet.Cells[1, findRange1.Column + 1], activeSheet.Cells[last.Row, last.Column + 1].EntireColumn];
									leftUsedCol.Interior.ColorIndex = 0;
									rightUsedCol.Interior.ColorIndex = 0;

									DialogResult dialogResult = MessageBoxEx.Show($"This will overwrite column {findRange1.Column}.\nAre you sure?", "Column Overwrite",
										MessageBoxButtons.YesNo, MessageBoxIcon.Warning, timeout);

									if (dialogResult == DialogResult.Yes)
									{
										for (int i = 3; i < numOfPoints; i++)
										{
											if (i > last.Row)
											{
												activeSheet.Cells[i, 6] = t.Item4.CurveData[i - 3].DoseValue.Dose;
											}

											activeSheet.Cells[i, findRange1.Column] = t.Item4.CurveData[i - 3].Volume;
											activeSheet.Cells[i, findRange1.Column].EntireColumn.Interior.Color = System.Drawing.Color.Yellow;

										}
									}
									else
									{
										goto nextPatient;
									}
								}
								else
								{
									Excel.Range range1Next1 = activeSheet.Range[activeSheet.Cells[2, findRange1.Column + 1], activeSheet.Cells[2, last.Column + 5]];
									findRange1Next1 = range1Next1.Find(utterance1);
									if (findRange1Next1 == null)
									{
										activeSheet.Cells[2, colCount].NumberFormat = "@";
										activeSheet.Cells[2, colCount].Font.Bold = true;

										activeSheet.Cells[1, colCount] = activePlanId;
										activeSheet.Cells[2, colCount] = studyId;

										for (int i = 3; i < numOfPoints; i++)
										{
											var lastUsedCol = activeSheet.Range[activeSheet.Cells[1, 6], activeSheet.Cells[last.Row, colCount - 1].EntireColumn];
											lastUsedCol.Interior.ColorIndex = 0;

											if (i > last.Row)
											{
												activeSheet.Cells[i, 6] = t.Item4.CurveData[i - 3].DoseValue.Dose;
											}

											activeSheet.Cells[i, colCount] = t.Item4.CurveData[i - 3].Volume;
											activeSheet.Cells[i, colCount].EntireColumn.Interior.Color = System.Drawing.Color.Yellow;
										}
									}
									else
									{
										string firstRowNext1 = activeSheet.Cells[1, findRange1Next1.Column].ToString();
										if (firstRowNext1 == utterance2)
										{
											if (findRange1Next1.Column < visibleRange.Column || findRange1Next1.Column > (visibleRange.Column + visibleRange.Columns.Count - 2))
											{
												// Scroll Excel worksheet if the row we are writing to is offscreen.
												int colToScrollTo = findRange1Next1.Column - visibleRange.Columns.Count;

												if (colToScrollTo < 1)
												{
													oXL.ActiveWindow.ScrollColumn = 1;
												}
												else
												{
													oXL.ActiveWindow.ScrollColumn = colToScrollTo + 6;
												}
											}

											var leftUsedCol = activeSheet.Range[activeSheet.Cells[1, 6], activeSheet.Cells[last.Row, findRange1Next1.Column - 1].EntireColumn];
											var rightUsedCol = activeSheet.Range[activeSheet.Cells[1, findRange1Next1.Column + 1], activeSheet.Cells[last.Row, last.Column + 1].EntireColumn];
											leftUsedCol.Interior.ColorIndex = 0;
											rightUsedCol.Interior.ColorIndex = 0;

											DialogResult dialogResult = MessageBoxEx.Show($"This will overwrite column {findRange1Next1.Column}.\nAre you sure?", "Column Overwrite",
												MessageBoxButtons.YesNo, MessageBoxIcon.Warning, timeout);

											if (dialogResult == DialogResult.Yes)
											{
												for (int i = 3; i < numOfPoints; i++)
												{
													if (i > last.Row)
													{
														activeSheet.Cells[i, 6] = t.Item4.CurveData[i - 3].DoseValue.Dose;
													}

													activeSheet.Cells[i, findRange1Next1.Column] = t.Item4.CurveData[i - 3].Volume;
													activeSheet.Cells[i, findRange1Next1.Column].EntireColumn.Interior.Color = System.Drawing.Color.Yellow;

												}
											}
											else
											{
												goto nextPatient;
											}
										}
										else
										{
											Excel.Range range1Next2 = activeSheet.Range[activeSheet.Cells[2, findRange1Next1.Column + 1], activeSheet.Cells[2, last.Column + 5]];
											findRange1Next2 = range1Next2.Find(utterance1);
											if (findRange1Next2 == null)
											{
												activeSheet.Cells[2, colCount].NumberFormat = "@";
												activeSheet.Cells[2, colCount].Font.Bold = true;

												activeSheet.Cells[1, colCount] = activePlanId;
												activeSheet.Cells[2, colCount] = studyId;

												for (int i = 3; i < numOfPoints; i++)
												{
													var lastUsedCol = activeSheet.Range[activeSheet.Cells[1, 6], activeSheet.Cells[last.Row, colCount - 1].EntireColumn];
													lastUsedCol.Interior.ColorIndex = 0;

													if (i > last.Row)
													{
														activeSheet.Cells[i, 6] = t.Item4.CurveData[i - 3].DoseValue.Dose;
													}

													activeSheet.Cells[i, colCount] = t.Item4.CurveData[i - 3].Volume;
													activeSheet.Cells[i, colCount].EntireColumn.Interior.Color = System.Drawing.Color.Yellow;
												}
											}
											else
											{
												string firstRowNext2 = activeSheet.Cells[1, findRange1Next2.Column].ToString();
												if (firstRowNext2 == utterance2)
												{
													if (findRange1Next2.Column < visibleRange.Column || findRange1Next2.Column > (visibleRange.Column + visibleRange.Columns.Count - 2))
													{
														// Scroll Excel worksheet if the row we are writing to is offscreen.
														int colToScrollTo = findRange1Next2.Column - visibleRange.Columns.Count;

														if (colToScrollTo < 1)
														{
															oXL.ActiveWindow.ScrollColumn = 1;
														}
														else
														{
															oXL.ActiveWindow.ScrollColumn = colToScrollTo + 6;
														}
													}

													var leftUsedCol = activeSheet.Range[activeSheet.Cells[1, 6], activeSheet.Cells[last.Row, findRange1Next2.Column - 1].EntireColumn];
													var rightUsedCol = activeSheet.Range[activeSheet.Cells[1, findRange1Next2.Column + 1], activeSheet.Cells[last.Row, last.Column + 1].EntireColumn];
													leftUsedCol.Interior.ColorIndex = 0;
													rightUsedCol.Interior.ColorIndex = 0;

													DialogResult dialogResult = MessageBoxEx.Show($"This will overwrite column {findRange1Next2.Column}.\nAre you sure?", "Column Overwrite",
														MessageBoxButtons.YesNo, MessageBoxIcon.Warning, timeout);

													if (dialogResult == DialogResult.Yes)
													{
														for (int i = 3; i < numOfPoints; i++)
														{
															if (i > last.Row)
															{
																activeSheet.Cells[i, 6] = t.Item4.CurveData[i - 3].DoseValue.Dose;
															}

															activeSheet.Cells[i, findRange1Next2.Column] = t.Item4.CurveData[i - 3].Volume;
															activeSheet.Cells[i, findRange1Next2.Column].EntireColumn.Interior.Color = System.Drawing.Color.Yellow;

														}
													}
													else
													{
														goto nextPatient;
													}
												}
												else
												{
													Excel.Range range1Next3 = activeSheet.Range[activeSheet.Cells[2, findRange1Next2.Column + 1], activeSheet.Cells[2, last.Column + 5]];
													findRange1Next3 = range1Next3.Find(utterance1);
													if (findRange1Next3 == null)
													{
														activeSheet.Cells[2, colCount].NumberFormat = "@";
														activeSheet.Cells[2, colCount].Font.Bold = true;

														activeSheet.Cells[1, colCount] = activePlanId;
														activeSheet.Cells[2, colCount] = studyId;

														for (int i = 3; i < numOfPoints; i++)
														{
															var lastUsedCol = activeSheet.Range[activeSheet.Cells[1, 6], activeSheet.Cells[last.Row, colCount - 1].EntireColumn];
															lastUsedCol.Interior.ColorIndex = 0;

															if (i > last.Row)
															{
																activeSheet.Cells[i, 6] = t.Item4.CurveData[i - 3].DoseValue.Dose;
															}

															activeSheet.Cells[i, colCount] = t.Item4.CurveData[i - 3].Volume;
															activeSheet.Cells[i, colCount].EntireColumn.Interior.Color = System.Drawing.Color.Yellow;
														}
													}
													else
													{
														string firstRowNext3 = activeSheet.Cells[1, findRange1Next3.Column].ToString();
														if (firstRowNext3 == utterance2)
														{
															if (findRange1Next3.Column < visibleRange.Column || findRange1Next3.Column > (visibleRange.Column + visibleRange.Columns.Count - 2))
															{
																// Scroll Excel worksheet if the row we are writing to is offscreen.
																int colToScrollTo = findRange1Next3.Column - visibleRange.Columns.Count;

																if (colToScrollTo < 1)
																{
																	oXL.ActiveWindow.ScrollColumn = 1;
																}
																else
																{
																	oXL.ActiveWindow.ScrollColumn = colToScrollTo + 6;
																}
															}

															var leftUsedCol = activeSheet.Range[activeSheet.Cells[1, 6], activeSheet.Cells[last.Row, findRange1Next3.Column - 1].EntireColumn];
															var rightUsedCol = activeSheet.Range[activeSheet.Cells[1, findRange1Next3.Column + 1], activeSheet.Cells[last.Row, last.Column + 1].EntireColumn];
															leftUsedCol.Interior.ColorIndex = 0;
															rightUsedCol.Interior.ColorIndex = 0;

															DialogResult dialogResult = MessageBoxEx.Show($"This will overwrite column {findRange1Next3.Column}.\nAre you sure?", "Column Overwrite",
																MessageBoxButtons.YesNo, MessageBoxIcon.Warning, timeout);

															if (dialogResult == DialogResult.Yes)
															{
																for (int i = 3; i < numOfPoints; i++)
																{
																	if (i > last.Row)
																	{
																		activeSheet.Cells[i, 6] = t.Item4.CurveData[i - 3].DoseValue.Dose;
																	}

																	activeSheet.Cells[i, findRange1Next3.Column] = t.Item4.CurveData[i - 3].Volume;
																	activeSheet.Cells[i, findRange1Next3.Column].EntireColumn.Interior.Color = System.Drawing.Color.Yellow;

																}
															}
															else
															{
																goto nextPatient;
															}
														}
														else
														{
															Excel.Range range1Next4 = activeSheet.Range[activeSheet.Cells[2, findRange1Next3.Column + 1], activeSheet.Cells[2, last.Column + 5]];
															findRange1Next4 = range1Next4.Find(utterance1);
															if (findRange1Next4 == null)
															{
																activeSheet.Cells[2, colCount].NumberFormat = "@";
																activeSheet.Cells[2, colCount].Font.Bold = true;

																activeSheet.Cells[1, colCount] = activePlanId;
																activeSheet.Cells[2, colCount] = studyId;

																for (int i = 3; i < numOfPoints; i++)
																{
																	var lastUsedCol = activeSheet.Range[activeSheet.Cells[1, 6], activeSheet.Cells[last.Row, colCount - 1].EntireColumn];
																	lastUsedCol.Interior.ColorIndex = 0;

																	if (i > last.Row)
																	{
																		activeSheet.Cells[i, 6] = t.Item4.CurveData[i - 3].DoseValue.Dose;
																	}

																	activeSheet.Cells[i, colCount] = t.Item4.CurveData[i - 3].Volume;
																	activeSheet.Cells[i, colCount].EntireColumn.Interior.Color = System.Drawing.Color.Yellow;
																}
															}
															else
															{
																string firstRowNext4 = activeSheet.Cells[1, findRange1Next4.Column].ToString();
																if (firstRowNext4 == utterance2)
																{
																	if (findRange1Next4.Column < visibleRange.Column || findRange1Next4.Column > (visibleRange.Column + visibleRange.Columns.Count - 2))
																	{
																		// Scroll Excel worksheet if the row we are writing to is offscreen.
																		int colToScrollTo = findRange1Next4.Column - visibleRange.Columns.Count;

																		if (colToScrollTo < 1)
																		{
																			oXL.ActiveWindow.ScrollColumn = 1;
																		}
																		else
																		{
																			oXL.ActiveWindow.ScrollColumn = colToScrollTo + 6;
																		}
																	}

																	var leftUsedCol = activeSheet.Range[activeSheet.Cells[1, 6], activeSheet.Cells[last.Row, findRange1Next4.Column - 1].EntireColumn];
																	var rightUsedCol = activeSheet.Range[activeSheet.Cells[1, findRange1Next4.Column + 1], activeSheet.Cells[last.Row, last.Column + 1].EntireColumn];
																	leftUsedCol.Interior.ColorIndex = 0;
																	rightUsedCol.Interior.ColorIndex = 0;

																	DialogResult dialogResult = MessageBoxEx.Show($"This will overwrite column {findRange1Next4.Column}.\nAre you sure?", "Column Overwrite",
																		MessageBoxButtons.YesNo, MessageBoxIcon.Warning, timeout);

																	if (dialogResult == DialogResult.Yes)
																	{
																		for (int i = 3; i < numOfPoints; i++)
																		{
																			if (i > last.Row)
																			{
																				activeSheet.Cells[i, 6] = t.Item4.CurveData[i - 3].DoseValue.Dose;
																			}

																			activeSheet.Cells[i, findRange1Next4.Column] = t.Item4.CurveData[i - 3].Volume;
																			activeSheet.Cells[i, findRange1Next4.Column].EntireColumn.Interior.Color = System.Drawing.Color.Yellow;

																		}
																	}
																	else
																	{
																		MessageBoxEx.Show("Got to Here somehow", timeout);
																		goto nextPatient;
																	}
																}
																else
																{

																}
															}
														}
													}
												}
											}
										}
									}
								}
							}
							TestCount++;
						}

					nextPatient:

						if ((saveCount)%10 == 0)
						{
							if (File.Exists(filePath))
							{
								oWB.Save();
							}
							else
							{
								oWB.SaveAs(filePath, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
									false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
									Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
							}
						}

						Console.Write($"\rCompleted DVH Extraction for {usableID}. Loading next patient...  ");
						MessageBoxEx.Show($"DVH queries successfully extracted for {patient.LastName}, {patient.FirstName} ({patient.Id})", "Extraction Complete", timeout);
					}
					else
					{
						//Self-exiting messagebox pop up (change timeout timer at top to change exit time)
						MessageBoxEx.Show($"Template does not exist for {usableID}.", "Extraction Failed",
										MessageBoxButtons.OK, MessageBoxIcon.Error, timeout);
						//if extraction fails writes to failed extractions w/ ID and reason
						using (System.IO.StreamWriter failedExtractionFile =
							new System.IO.StreamWriter(failedExtractionLoc, true))
						{
							failedExtractionFile.Write(usableID);
							failedExtractionFile.WriteLine(" -- Template did not exist in Patient Template Location.");
						}
					}
					saveCount++;
				}
				for (int listCount = 0; listCount < subStruc.Count; listCount++)
				{
					string sheetPick = subStruc[listCount];

					var usingSheet = oWB.Sheets[$"{sheetPick}"];
					usingSheet.Select();

					Excel.Range last = usingSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
					string columnLetter = ColumnIndexToColumnLetter(last.Column);

                    if (usingSheet.Cells[last.Row + 1, 1] != null)
                    {
                        usingSheet.Range($"G1:{columnLetter}{last.Row}").ColumnWidth = 10.00;
                        usingSheet.Cells[3, 1] = $"=AVERAGE(G3:{columnLetter}3)";
                        usingSheet.Cells[3, 2] = $"=MIN(G3:{columnLetter}3)";
                        usingSheet.Cells[3, 3] = $"=MAX(G3:{columnLetter}3)";
                        usingSheet.Cells[3, 4] = $"=STDEV(G3:{columnLetter}3)";
                    }

					oXL.ActiveWindow.ScrollColumn = 1;

					for (int j = 1; j <= 4; j++)
					{
						if (usingSheet.Cells[3, 7] != null)
						{
							string colLet = ColumnIndexToColumnLetter(j);
							Excel.Range range = usingSheet.Range($"{colLet}3:{colLet}3", Type.Missing);
							Excel.Range dest = null;
							if (last.Row > 3)
							{
								dest = usingSheet.Range($"{colLet}3:{colLet}{last.Row}");
							}
							else
							{
								dest = usingSheet.Range($"{colLet}3:{colLet}10");
							}
							range.AutoFill(dest, Excel.XlAutoFillType.xlFillCopy);
						}
					}
                    
                    Excel.Range chartRange;
                    Excel.ChartObjects xlCharts = (Excel.ChartObjects)usingSheet.ChartObjects(Type.Missing);
                    Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(10, 80, 400, 332);
                    myChart.Name = "exChart";
                    Excel.Chart chartPage = myChart.Chart;

                    if (last.Column < 255)
                    {
                        chartRange = usingSheet.Range("F2", $"{columnLetter}{last.Row}");
                        chartPage.SetSourceData(chartRange, misvalue);
                        chartPage.ChartType = Excel.XlChartType.xlXYScatterSmoothNoMarkers;
                        chartPage.Legend.Clear();
                        chartPage.HasTitle = false;
                    }
                    else
                    {
                        MessageBoxEx.Show($"Maximum Number of Data Series in chart is exceeded.\nLimiting chart up to {usingSheet.Cells[2, 255].Value2.ToString()}.", "Data Series Limit Exceeded", MessageBoxButtons.OK, MessageBoxIcon.Warning, timeout);
                        using (System.IO.StreamWriter failedExtractionFile =
                            new System.IO.StreamWriter(failedExtractionLoc, true))
                        {
                            failedExtractionFile.Write(usableID);
                            failedExtractionFile.WriteLine($" -- Maximum number of data series in chart was exceeded for Site: {structure}, Sheet:{usingSheet.Name.ToString()}");
                        }
                        columnLetter = ColumnIndexToColumnLetter(255);
                        chartRange = usingSheet.Range("F2", $"{columnLetter}{last.Row}");
                        chartPage.SetSourceData(chartRange, misvalue);
                        chartPage.ChartType = Excel.XlChartType.xlXYScatterSmoothNoMarkers;
                        chartPage.Legend.Clear();
                        chartPage.HasTitle = false;
                    }
				}

                for (int i = oXL.ActiveWorkbook.Worksheets.Count; i > 0; i--)
                {
                    Excel.Worksheet wKSheet = (Excel.Worksheet)oXL.ActiveWorkbook.Worksheets[i];
                    if (wKSheet.Name == "Sheet1")
                    {
                        wKSheet.Delete();
                    }
                }

                Excel.Worksheet firstSheet = oWB.Sheets[1];
                firstSheet.Select();

				if (File.Exists(filePath))
				{
					oWB.Save();
				}
				else
				{
					oWB.SaveAs(filePath, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
						false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
						Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
				}

				DateTime localDateB = DateTime.Now;
				string localTimeStringB = localDateB.ToString("hh:mm:ss tt");
				string localDateStringB = localDateB.ToString("yyyy/MM/dd hh:mm:ss tt");
				Console.WriteLine($"Completed: {localTimeStringB}");

				MessageBoxEx.Show($"{structure}" +
					$"\nStarted: {localTimeStringA}" +
					$"\nCompleted: {localTimeStringB}", "Start - End", 15000);

				using (System.IO.StreamWriter failedExtractionFile =
					new System.IO.StreamWriter(failedExtractionLoc, true))
				{
					failedExtractionFile.Write($"Completed Extraction: {localDateStringB}");
				}

				Console.WriteLine("Loading Failed Extractions.");
				System.Diagnostics.Process.Start(failedExtractionLoc);
				//oWB.Close();
				//oXL.Quit();
			}
			else
			{
				MessageBox.Show("Patient Template for have not been created yet.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

        // Attempts to convert a string to an integer.
        // Returns 0 if the conversion fails.
        static int StringToInt(string s)
        {
            int i = 0;
            if (!Int32.TryParse(s, out i))
            {
                i = 0;
                //System.Diagnostics.Debug.Assert(i != -1);
            }
            return i;
        }

        // Checks if two strings are the same (case-insensitive).
        static bool CheckIdMatch(string s1, string s2)
        {
            bool result = s1.Equals(s2, StringComparison.OrdinalIgnoreCase);
            return result;
        }

        // Removes characters from a string taht are not letters.
        static string RemoveSpecialChar(string s)
        {
            StringBuilder sb = new StringBuilder();
            foreach (char c in s)
            {
                if ((c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z'))
                {
                    sb.Append(c);
                }
            }
            return sb.ToString();
        }

        // Generates a patient's study ID using the default conventions.
        static string GenerateStudyId(Patient patient)
        {
            if (patient == null)
            {
                return "";
            }

            string lastName = patient.LastName;
            string patientId = patient.Id;

            lastName = GetSubstringByLength(lastName, 0, 3);
            lastName = CharFiller(lastName, '_', 3);
            patientId = GetSubstringByLength(patientId, 2, 4);
            patientId = CharFiller(patientId, '_', 4);

            return (lastName + patientId).ToUpper();
        }

        // Returns a substring from index startIndex of length maxLen or less, depending on
        // the number of characters in the substring.
        public static string GetSubstringByLength(string str, int startIndex, int maxLen)
        {
            if (string.IsNullOrEmpty(str))
            {
                return str;
            }
            return str.Substring(startIndex, Math.Min(str.Length - startIndex, maxLen));
        }

        // Concatenates character c onto string str until the string length is equal to len.
        public static string CharFiller(string str, char c, int len)
        {
            if (string.IsNullOrEmpty(str))
            {
                return str;
            }

            int strLen = str.Length;
            if (strLen >= len)
            {
                return str;
            }

            StringBuilder sb = new StringBuilder();
            sb.Append(str);
            for (int i = strLen; i < len; i++)
            {
                sb.Append(c);
            }
            return sb.ToString();
        }

        // Returns true if an IEnumerable<T> is not null and not empty.
        static bool IsNullOrEmpty<T>(IEnumerable<T> data)
        {
            return data == null || !data.Any();
        }

        // Reads in lines while ignoring blank/empty lines.
        static string[] ReadNonBlankLines(string path)
        {
            string line;
            List<string> lines = new List<string>();

            using (StreamReader sr = new StreamReader(path))
                while (true)
                {
                    line = sr.ReadLine();

                    if (line == null)
                    {
                        break;
                    }
                    else if (line == Environment.NewLine || line == "")
                    {
                        continue;
                    }
                    lines.Add(line.Trim());
                }

            return lines.ToArray();
        }

        static string MakeSafeFilename(string filename)
        {
            //return string.Join("_", filename.Split(Path.GetInvalidFileNameChars()));
            return string.Join("_", filename.Split(new char[] { '/', '+', ':', ',', '.' }));
        }

		public class Values
		{
			public string inputInfo { get; set; }
			public string structure { get; set; }

			public void getStructure()
			{
				string workingLocExternal = @"\\dc3-pr-files\MedPhysics Backup\Data Extractions\Automation\Info Transfer External.txt";

				structure = "";
				//if (new FileInfo(workingLocExternal).Length != 0)
				using (StreamReader sr = new StreamReader(workingLocExternal))
				{
					string contents = sr.ReadToEnd();
					if (!contents.Contains("\\") && new FileInfo(workingLocExternal).Length != 0)
					{
						structure = File.ReadLines(workingLocExternal).Skip(0).Take(1).First();
						DialogResult result = MessageBoxEx.Show($"Would you like to run with {structure} as the set structure for DVH?", "DVH Input Required",
							MessageBoxButtons.YesNo, MessageBoxIcon.Question, 7500);
						if (result == DialogResult.Yes)
						{
							structure = File.ReadLines(workingLocExternal).Skip(0).Take(1).First();
						}
						else if (result == DialogResult.No)
						{
							structure = "";
						}
						else
						{
							structure = File.ReadLines(workingLocExternal).Skip(0).Take(1).First();
						}
					}
				}
			}

			public void getInputInfo()
			{
				inputInfo = $@"\\dc3-pr-files\MedPhysics Backup\Data Extractions\Automation\DVH\Input Information Text Files\{structure} DVH Input Information.txt";
			}
		}

		static string ColumnIndexToColumnLetter(int colIndex)
		{
			int div = colIndex;
			string colLetter = String.Empty;
			int mod = 0;

			while (div > 0)
			{
				mod = (div - 1) % 26;
				colLetter = (char)(65 + mod) + colLetter;
				div = (int)((div - mod) / 26);
			}
			return colLetter;
		}
	}
}