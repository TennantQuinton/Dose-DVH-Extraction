using System;
using System.Linq;
using System.Text;
using System.Collections.Generic;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;
using System.Windows.Forms;
using System.IO;
using Application = VMS.TPS.Common.Model.API.Application;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace TestDoseExtraction
{
    public class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            uint timeout = 0; //messagebox timeout timer
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
            string workingLocExternal = @"\\dc3-pr-files\MedPhysics Backup\Data Extractions\Automation\Info Transfer External.txt";
			string workingLocInternal = @"\\dc3-pr-files\MedPhysics Backup\Data Extractions\Automation\Info Transfer Internal.txt";

			Console.WriteLine("Last Updated: 2019-12-17 by Tennant, Quinton");

            //Get todays date to identify when the failed extraction document was made
            DateTime localDate = DateTime.Now;
            string localDateString = localDate.ToString("yyyy/MM/dd hh:mm:ss tt");

            DateTime localDateA = DateTime.Now;
            string localTimeStringA = localDateA.ToString("hh:mm:ss tt");
            Console.WriteLine("Started: "+ localTimeStringA);

            string structure = "";

			//if (new FileInfo(workingLocExternal).Length != 0)
			using (StreamReader sr = new StreamReader(workingLocExternal))
			{
				string contents = sr.ReadToEnd();
				if (!contents.Contains("\\") && new FileInfo(workingLocExternal).Length != 0)
				{
					structure = File.ReadLines(workingLocExternal).Skip(0).Take(1).First();
					DialogResult result = MessageBoxEx.Show($"Would you like to run with {structure} as the set structure for Dose?", "Dose Input Required",
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

            string inputInfo;
            inputInfo = $@"\\dc3-pr-files\MedPhysics Backup\Data Extractions\Automation\Dose\Input Information Text Files\{structure} Dose Input Information.txt";

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

            //string inputInfo = @"\\\dc3-pr-files\MedPhysics Backup\Coop Students\2019\Term 3 - Autumn\Quinton Tennant\Input Information.txt";
            //Getting input and output information from input information text file so users don't have to enter coding env
            //Working file is used as a roundabout way to pass patient ID DVHQuery.cs (if you can find a better way, go ahead!)
            string failedExtractionLoc = File.ReadLines(inputInfo).Skip(36).Take(1).First();
            string usableIDLoc = File.ReadLines(inputInfo).Skip(32).Take(1).First();
            string textFileLoc = File.ReadLines(inputInfo).Skip(16).Take(1).First();
            string usableID;

            uint timeout = 0; //messagebox timeout timer
            int count = 0;

            //Empties working file and writes the created date header in failed extractions
            File.WriteAllText(@failedExtractionLoc, "Created "+ localDateString + System.Environment.NewLine);

            if (File.Exists(usableIDLoc))
            {
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

                    ExcelExtensions.ExcelHelper excelHelper = new ExcelExtensions.ExcelHelper(patient);
                    // Check if the user has Excel installed, which is required to use
                    // the Excel Interop library.
                    excelHelper.CheckInstallation();

                    //creates template text files for each ID during the loop
                    string textFilePath = $@"{textFileLoc}\{usableID}.txt";

                    if (File.Exists(textFilePath))
                    {
						Console.Write($"\rRunning Dose Extraction on {usableID}                                  ");
                        // We use a list PlanQueries to separate DVHQueries by their plan. This makes it easier to write into
                        // Excel, because each patient plan will go into a separate worksheet.
                        List<PlanQueries> queryList = new List<PlanQueries>();

                        // Used for storing the plan names and structures which could not be found.
                        List<Tuple<string, string>> errorList = new List<Tuple<string, string>>();

                        // Reads each line of the file into a string array, ignoring empty lines.
                        string[] lines = IOExtensions.ReadNonBlankLines(textFilePath);

                        // Check for invalid characters in the save file path.
                        if (IOExtensions.CheckForInvalidPathChars(lines[0]))
                        {
                            MessageBoxEx.Show("You have invalid characters in your file path:" + Environment.NewLine
                                + "\t* ? \" < > |", "Invalid File Path",
                                                    MessageBoxButtons.OK, MessageBoxIcon.Error, timeout);
                            return;
                        }

                        //string excelFilename = Path.GetFileName(lines[0]);
						//string saveDirectory = File.ReadLines(inputInfo).Skip(40).Take(1).First();
						//
						//saveDirectory = EclipseExtensions.FormatCitrixSaveDirectory(saveDirectory);
						//
                        string filePath = File.ReadLines(inputInfo).Skip(40).Take(1).First();

						excelHelper.SetFilePath(filePath);

                        // Try creating the save directory. If it already exists, nothing happens.
                        //Directory.CreateDirectory(saveDirectory);

                        int index = 1; // lines[index] == Course
                        int numOfLines = lines.Length;

                        // Check text file for input errors.
                        int numOfStructs;
                        int structIndex = 3;

                        while (structIndex < numOfLines)
                        {
                            if (Int32.TryParse(lines[structIndex], out numOfStructs))
                            {
                                structIndex = structIndex + numOfStructs + 3;
                            }
                            else
                            {
                                //Self-exiting messagebox pop up (change timeout timer at top to change exit time)
                                MessageBoxEx.Show($"The number of structures in the text file does not match the inputted number for one of the plans.", "Extraction Failed",
                                                MessageBoxButtons.OK, MessageBoxIcon.Error, timeout);
                                //if extraction fails writes to failed extractions w/ ID and reason
                                using (System.IO.StreamWriter failedExtractionFile =
                                    new System.IO.StreamWriter(failedExtractionLoc, true))
                                {
                                    failedExtractionFile.Write(usableID);
                                    failedExtractionFile.WriteLine(" -- Number of structures in the text file did not match inputted number for one of the plans.");
                                }
                                //instead of exiting the code on fail we want to go to end of this loop to begin next
                                goto failedException;
                            }
                        }

                        for (; index < numOfLines;)
                        {
                            // Check if the Patient has Courses. If not, there is no DVH data to extract
                            // so we return.
                            if (IsNullOrEmpty(patient.Courses))
                            {
                                //Self-exiting messagebox pop up (change timeout timer at top to change exit time)
                                MessageBoxEx.Show($"{patient.LastName}, {patient.FirstName} ({patient.Id}) does not have any courses.", "Extraction Failed",
                                                MessageBoxButtons.OK, MessageBoxIcon.Error, timeout);
                                //if extraction fails writes to failed extractions w/ ID and reason
                                using (System.IO.StreamWriter failedExtractionFile =
                                    new System.IO.StreamWriter(failedExtractionLoc, true))
                                {
                                    failedExtractionFile.Write(usableID);
                                    failedExtractionFile.WriteLine(" -- Patient does not have any courses.");
                                }
                                //do not exit code, go to next loop iteration
                                goto failedException;
                            }
                            Course course = EclipseExtensions.GetCourseByID(patient, lines[index]);
                            index++; // lines[index] == Excel Plan Name : Eclipse Plan ID

                            if (course == null)
                            {
                                //Self-exiting messagebox pop up (change timeout timer at top to change exit time)
                                DialogResult dialogResult = MessageBoxEx.Show($"Course {lines[index - 1]} could not be found. Would you like to continue?", "Warning",
                                                                            MessageBoxButtons.OKCancel, MessageBoxIcon.Warning, timeout);
                                //write to failed extraction w/ ID and reason
                                using (System.IO.StreamWriter failedExtractionFile =
                                    new System.IO.StreamWriter(failedExtractionLoc, true))
                                {
                                    failedExtractionFile.Write(usableID);
                                    failedExtractionFile.WriteLine($" -- Dose extraction for course {lines[index - 1]} could not be found.");
                                }

                                if (dialogResult == DialogResult.Cancel)
                                {
                                    //Self-exiting messagebox pop up (change timeout timer at top to change exit time)
                                    MessageBoxEx.Show($"Dose extraction could not be completed for {patient.LastName}, {patient.FirstName} ({patient.Id}).", "Extraction Failed",
                                                        MessageBoxButtons.OK, MessageBoxIcon.Error, timeout);
                                    //write to failed extraction w/ ID and reason
                                    using (System.IO.StreamWriter failedExtractionFile =
                                        new System.IO.StreamWriter(failedExtractionLoc, true))
                                    {
                                        failedExtractionFile.Write(usableID);
                                        failedExtractionFile.WriteLine($" -- Cancelled Dose extraction after course {lines[index - 1]} could not be found.");
                                    }
                                    //do not exit code, go to next loop iteration
                                    goto failedException;
                                }
                                // Increments the index so that on the next loop we are immediately at the next Course.
                                // (Skips looking at the plans and structures).
                                index = index + Convert.ToInt32(lines[index + 1]) + 2;
                                continue;
                            }
                            else
                            {
                                // Check if the Course has any PlanSetups/PlanSums. If not, a different Course
                                // could have plans, so we continue looping until we get to the next Course in
                                // the text file.
                                if (IsNullOrEmpty(course.PlanSetups) && IsNullOrEmpty(course.PlanSums))
                                {
                                    //Self-exiting messagebox pop up (change timeout timer at top to change exit time)
                                    DialogResult dialogResult = MessageBoxEx.Show($"Course {lines[index - 1]} does not contain any plans. Would you like to continue?", "Warning",
                                                                                MessageBoxButtons.OKCancel, MessageBoxIcon.Warning, timeout);
                                    //write to failed extraction w/ ID and reason
                                    using (System.IO.StreamWriter failedExtractionFile =
                                        new System.IO.StreamWriter(failedExtractionLoc, true))
                                    {
                                        failedExtractionFile.Write(usableID);
                                        failedExtractionFile.WriteLine($" -- Course {lines[index - 1]} did not contain any plans.");
                                    }

                                    if (dialogResult == DialogResult.Cancel)
                                    {
                                        //Self-exiting messagebox pop up (change timeout timer at top to change exit time)
                                        MessageBoxEx.Show($"Dose extraction could not be completed for {patient.LastName}, {patient.FirstName} ({patient.Id}).", "Extraction Failed",
                                                        MessageBoxButtons.OK, MessageBoxIcon.Error, timeout);

                                        using (System.IO.StreamWriter failedExtractionFile =
                                            new System.IO.StreamWriter(failedExtractionLoc, true))
                                        {
                                            //write to failed extraction w/ ID and reason
                                            failedExtractionFile.Write(usableID);
                                            failedExtractionFile.WriteLine($" -- Cancelled after course {lines[index - 1]} did not contain any plans.");
                                        }
                                        //do not exit code, go to next loop iteration, do not collect $200
                                        goto failedException;
                                    }
                                    // Increments the index so that on the next loop we are immediately at the next Course.
                                    // (Skips looking at the plans and structures).
                                    index = index + Convert.ToInt32(lines[index + 1]) + 2;
                                    continue;
                                }

                                // Adds the Plan to the Excel ID Dictionary, using the actual Plan ID in Eclipse
                                // as the key and the desired Plan name (worksheet name) in Excel as the value.
                                string[] planLine = lines[index].Split('|');
                                string eclipsePlanID = planLine[1].Trim().ToUpper();
                                string excelPlanName = planLine[0].Trim();

                                // Store the total prescribed dose, so that the same value can be used for Plan Setups and Plan Sums.
                                DoseValue? totalPrescribedDose = null;

                                if (planLine.Length > 2)
                                {
                                    try
                                    {
                                        totalPrescribedDose = EclipseExtensions.ParsePrescribedDose(planLine[2].Trim());
                                    }
                                    catch (Exception e)
                                    {
                                        //Self-exiting messagebox pop up (change timeout timer at top to change exit time)
                                        DialogResult dialogResult = MessageBoxEx.Show($"No total prescribed dose for the Plan {eclipsePlanID} was inputted. Would you like to continue?", "Warning",
                                                                                    MessageBoxButtons.OKCancel, MessageBoxIcon.Warning, timeout);
                                        //write to failed extraction w/ ID and reason
                                        using (System.IO.StreamWriter failedExtractionFile =
                                            new System.IO.StreamWriter(failedExtractionLoc, true))
                                        {
                                            failedExtractionFile.Write(usableID);
                                            failedExtractionFile.WriteLine($" -- No total prescribed dose for the Plan, {eclipsePlanID}, was inputted.");
                                        }

                                        if (dialogResult == DialogResult.Cancel)
                                        {
                                            //Self-exiting messagebox pop up (change timeout timer at top to change exit time)
                                            MessageBoxEx.Show($"Dose extraction could not be completed for {patient.LastName}, {patient.FirstName} ({patient.Id}).", "Extraction Failed",
                                                        MessageBoxButtons.OK, MessageBoxIcon.Error, timeout);

                                            //write to failed extraction w/ ID and reason
                                            using (System.IO.StreamWriter failedExtractionFile =
                                                new System.IO.StreamWriter(failedExtractionLoc, true))
                                            {
                                                failedExtractionFile.Write(usableID);
                                                failedExtractionFile.WriteLine($" -- Cancelled after no total prescribed dose for the Plan, {eclipsePlanID}, was inputted.");
                                            }
                                            //do not exit code, go to next loop iteration
                                            goto failedException;
                                        }
                                    }

                                }

                                List<string> planSumPlanIDs = new List<string>();
                                int planLineLen = planLine.Length;

                                // Add Plan Setup IDs from text file to a list.
                                for (int i = 3; i < planLineLen; i++)
                                {
                                    if (string.IsNullOrWhiteSpace(planLine[i]))
                                    {
                                        planSumPlanIDs.Add("");
                                    }
                                    else
                                    {
                                        planSumPlanIDs.Add(planLine[i].Trim().ToUpper());
                                    }
                                }

                                excelHelper.AddToExcelIDs(eclipsePlanID, excelPlanName);

                                PlanSetup planSetup = EclipseExtensions.GetPlanSetupByID(course, eclipsePlanID);
                                PlanSum planSum = EclipseExtensions.GetPlanSumByID(course, eclipsePlanID);
                                index++; // lines[index] == Number of Structure Queries

                                double ptvVolume = Double.NaN;

                                if (planSetup != null)
                                {
                                    totalPrescribedDose = planSetup.TotalPrescribedDose;

                                    // Check if the PlanSetup has a StructureSet/Structures. If not, a
                                    // different plan could have them, so we continue looping until we get
                                    // to the next Course in the text file (could be the same Course).
                                    if (planSetup.StructureSet == null || IsNullOrEmpty(planSetup.StructureSet.Structures))
                                    {
                                        //Self-exiting messagebox pop up (change timeout timer at top to change exit time)
                                        DialogResult dialogResult = MessageBoxEx.Show($"Plan {eclipsePlanID} does not contain any structures. Would you like to continue?", "Warning",
                                                                                    MessageBoxButtons.OKCancel, MessageBoxIcon.Warning, timeout);
                                        //write in failed extraction with ID and reason
                                        using (System.IO.StreamWriter failedExtractionFile =
                                            new System.IO.StreamWriter(failedExtractionLoc, true))
                                        {
                                            failedExtractionFile.Write(usableID);
                                            failedExtractionFile.WriteLine($" -- Plan, {eclipsePlanID}, did not contain any structures.");
                                        }

                                        if (dialogResult == DialogResult.Cancel)
                                        {
                                            //Self-exiting messagebox pop up (change timeout timer at top to change exit time)
                                            MessageBoxEx.Show($"Dose extraction could not be completed for {patient.LastName}, {patient.FirstName} ({patient.Id}).", "Extraction Failed",
                                                        MessageBoxButtons.OK, MessageBoxIcon.Error, timeout);

                                            //write in failed extraction with ID and reason
                                            using (System.IO.StreamWriter failedExtractionFile =
                                                new System.IO.StreamWriter(failedExtractionLoc, true))
                                            {
                                                failedExtractionFile.Write(usableID);
                                                failedExtractionFile.WriteLine($" -- Cancelled after plan, {eclipsePlanID}, did not contain any structures.");
                                            }
                                            //do not exit code, go to next loop iteration
                                            goto failedException;
                                        }
                                        // Increments the index so that on the next loop we are immediately at the next Course.
                                        // (Skips looking at the structures).
                                        index = index + Convert.ToInt32(lines[index]) + 1;
                                        continue;
                                    }

                                    PlanQueries queriesByPlan = new PlanQueries(patient, course, planSetup, totalPrescribedDose);
                                    queryList.Add(queriesByPlan);

                                    int numOfStructures = Convert.ToInt32(lines[index]);
                                    index++; // lines[index] == Excel Structure Name : Eclipse Structure ID : Queries

                                    // Iterate through the structure set to try and find the PTV.
                                    for (int structCount = 0; structCount < numOfStructures; structCount++)
                                    {
                                        string[] structLine = lines[index + structCount].Split('|');
                                        string eclipseStructID = structLine[1].Trim().ToUpper();
                                        string excelStructName = structLine[0].Trim();

                                        Structure target = EclipseExtensions.GetStructureByID(planSetup, eclipseStructID);

                                        if (target == null)
                                        {
                                            continue;
                                        }

                                        // TO-DO: Let the user specify what target volume (i.e. PTV, CTV, GTV, ITV) they want
                                        //        to use for R_## queries.
                                        // Get PTV volume based on what is labelled as "PTV" in the text file.
                                        if (excelStructName.ToUpper().Contains("PTV"))
                                        {
                                            ptvVolume = target.Volume;
                                            break;
                                        }
                                    }

                                    for (int structCount = 0; structCount < numOfStructures; structCount++)
                                    {
                                        string[] structLine = lines[index].Split('|');
                                        string eclipseStructID = structLine[1].Trim().ToUpper();
                                        string excelStructName = structLine[0].Trim();

                                        string[] structQueries = structLine[2].Trim().Split(new char[] { ' ', '\t' },
                                                                                            StringSplitOptions.RemoveEmptyEntries);

                                        Structure target = EclipseExtensions.GetStructureByID(planSetup, eclipseStructID);

                                        // Add any structures that couldn't be found to the error output.
                                        if (target == null)
                                        {
                                            errorList.Add(new Tuple<string, string>(planSetup.Id, eclipseStructID));
                                        }

                                        // We store the structure's computed DVH so we don't need
                                        // to recompute it each time. We store the DVHData in the
                                        // for-loop and not in a dictionary because no structure
                                        // should appear twice in our query.
                                        DVHData structDvh = null;

                                        // Decreasing the step size results in more points in the DVH curve data, and although it is
                                        // more accurate, it also increases the computing time. 
                                        // The max, min, and mean doses are unaffected by the step size.
                                        // The default step size Varian uses in calculating the DVH is 0.001.
                                        double stepSize = 0.01;

                                        // Calculate the DVHData for the structure.
                                        if (IsNullOrEmpty(structQueries))
                                        {
                                            // If we only need the max, min and mean doses, we calculate an inaccurate DVH.
                                            structDvh = planSetup.GetDVHCumulativeData(target, DoseValuePresentation.Absolute,
                                                VolumePresentation.AbsoluteCm3, 0.1);
                                        }
                                        else
                                        {
                                            // If we need to calculate other doses, we calculate an accurate DVH.
                                            structDvh = planSetup.GetDVHCumulativeData(target, DoseValuePresentation.Absolute,
                                                VolumePresentation.AbsoluteCm3, stepSize);
                                        }

                                        // Every structure should by default output the max, min, and mean doses.
                                        queriesByPlan.AddDVHQuery(new DVHQuery("svol", planSetup, target, totalPrescribedDose, excelStructName, ptvVolume, structDvh));
                                        queriesByPlan.AddDVHQuery(new DVHQuery("dmax", planSetup, target, totalPrescribedDose, excelStructName, ptvVolume, structDvh));
                                        queriesByPlan.AddDVHQuery(new DVHQuery("dmin", planSetup, target, totalPrescribedDose, excelStructName, ptvVolume, structDvh));
                                        queriesByPlan.AddDVHQuery(new DVHQuery("dmean", planSetup, target, totalPrescribedDose, excelStructName, ptvVolume, structDvh));

                                        // Output the equivalent sphere diameter for specific target volumes.
                                        if (excelStructName.ToUpper().Contains("CTV") || eclipseStructID.Contains("CTV") ||
                                            excelStructName.ToUpper().Contains("GTV") || eclipseStructID.Contains("GTV") ||
                                            excelStructName.ToUpper().Contains("PTV") || eclipseStructID.Contains("PTV"))
                                        {
                                            queriesByPlan.AddDVHQuery(new DVHQuery("eqsd", planSetup, target, totalPrescribedDose, excelStructName, ptvVolume, structDvh));
                                        }

                                        foreach (string query in structQueries)
                                        {
                                            DVHQuery currentQuery = new DVHQuery(query, planSetup, target, totalPrescribedDose, excelStructName, ptvVolume, structDvh);
                                            queriesByPlan.AddDVHQuery(currentQuery);
                                        }

                                        index++; // lines[index] == Excel Structure Name : Eclipse Structure ID : Queries || Course
                                    }
                                }
                                else if (planSum != null)
                                {
                                    // Check if the PlanSum has a StructureSet/Structures. If not, a
                                    // different plan could have them, so we continue looping until we get
                                    // to the next Course in the text file (could be the same Course).
                                    if (planSum.StructureSet == null || IsNullOrEmpty(planSum.StructureSet.Structures))
                                    {
                                        //Self-exiting messagebox pop up (change timeout timer at top to change exit time)
                                        DialogResult dialogResult = MessageBoxEx.Show($"Plan {eclipsePlanID} does not contain any structures. Would you like to continue?", "Warning",
                                                                                    MessageBoxButtons.OKCancel, MessageBoxIcon.Warning, timeout);
                                        //write to failed extractions with ID and reason
                                        using (System.IO.StreamWriter failedExtractionFile =
                                            new System.IO.StreamWriter(failedExtractionLoc, true))
                                        {
                                            failedExtractionFile.Write(usableID);
                                            failedExtractionFile.WriteLine($" -- Plan, {eclipsePlanID}, did not contain any structures.");
                                        }

                                        if (dialogResult == DialogResult.Cancel)
                                        {
                                            //Self-exiting messagebox pop up (change timeout timer at top to change exit time)
                                            MessageBoxEx.Show($"Dose extraction could not be completed for {patient.LastName}, {patient.FirstName} ({patient.Id}).", "Extraction Failed",
                                                        MessageBoxButtons.OK, MessageBoxIcon.Error, timeout);
                                            //write to failed extractions with ID and reason
                                            using (System.IO.StreamWriter failedExtractionFile =
                                                new System.IO.StreamWriter(failedExtractionLoc, true))
                                            {
                                                failedExtractionFile.Write(usableID);
                                                failedExtractionFile.WriteLine($" -- Cancelled after Plan, {eclipsePlanID}, did not contain any structures.");
                                            }
                                            //do not exit code, go to next loop iteration
                                            goto failedException;
                                        }
                                        // Increments the index so that on the next loop we are immediately at the next Course.
                                        // (Skips looking at the structures).
                                        index = index + Convert.ToInt32(lines[index]) + 1;
                                        continue;
                                    }

                                    PlanQueries queriesByPlan = new PlanQueries(patient, course, planSum, totalPrescribedDose);
                                    queryList.Add(queriesByPlan);

                                    // Try to get the actual Eclipse structures from the inputted strings.
                                    foreach (string planID in planSumPlanIDs)
                                    {
                                        PlanSetup planSumPlan = null;

                                        foreach (Course planSumPlanCourse in patient.Courses)
                                        {
                                            planSumPlan = EclipseExtensions.GetPlanSetupByID(planSumPlanCourse, planID);
                                            if (planSumPlan != null)
                                            {
                                                queriesByPlan.AddPlanSumPlan(new PlanQueries(patient, planSumPlanCourse, planSumPlan, totalPrescribedDose));
                                                break;
                                            }
                                        }

                                        // Plan Setup could be null, but we still add it to the list so that
                                        // we can leave blank spaces in the Excel file.
                                        if (!string.IsNullOrWhiteSpace(planID) && planSumPlan == null)
                                        {
                                            //Self-exiting messagebox pop up (change timeout timer at top to change exit time)
                                            DialogResult planDialog = MessageBoxEx.Show($"Plan {planID} could not be found in the plan sum {planSum.Id}. Would you like to continue?", "Warning",
                                                                              MessageBoxButtons.OKCancel, MessageBoxIcon.Warning, timeout);
                                            //write to failed extractions with ID and reason
                                            using (System.IO.StreamWriter failedExtractionFile =
                                                new System.IO.StreamWriter(failedExtractionLoc, true))
                                            {
                                                failedExtractionFile.Write(usableID);
                                                failedExtractionFile.WriteLine($" -- Plan, {planID}, could not be found in the plan sum, {planSum.Id}.");
                                            }

                                            if (planDialog == DialogResult.Cancel)
                                            {
                                                //Self-exiting messagebox pop up (change timeout timer at top to change exit time)
                                                MessageBoxEx.Show($"Dose extraction could not be completed for {patient.LastName}, {patient.FirstName} ({patient.Id}).", "Extraction Failed",
                                                                MessageBoxButtons.OK, MessageBoxIcon.Error, timeout);
                                                //write to failed extractions with ID and reason
                                                using (System.IO.StreamWriter failedExtractionFile =
                                                    new System.IO.StreamWriter(failedExtractionLoc, true))
                                                {
                                                    failedExtractionFile.Write(usableID);
                                                    failedExtractionFile.WriteLine($" -- Cancelled after Plan, {planID}, could not be found in the plan sum, {planSum.Id}.");
                                                }
                                                //do not exit code, go to next loop iteration
                                                goto failedException;
                                            }
                                        }

                                        if (planSumPlan == null)
                                        {
                                            queriesByPlan.AddPlanSumPlan(new PlanQueries(patient, null, null, totalPrescribedDose));
                                        }
                                    }

                                    int numOfStructures = Convert.ToInt32(lines[index]);
                                    index++; // lines[index] == Excel Structure Name : Eclipse Structure ID : Queries

                                    // Iterate through the structure set to try and find the PTV.
                                    for (int structCount = 0; structCount < numOfStructures; structCount++)
                                    {
                                        string[] structLine = lines[index + structCount].Split('|');
                                        string eclipseStructID = structLine[1].Trim().ToUpper();
                                        string excelStructName = structLine[0].Trim();

                                        Structure target = EclipseExtensions.GetStructureByID(planSum, eclipseStructID);

                                        if (target == null)
                                        {
                                            continue;
                                        }

                                        // Get PTV volume based on what is labelled as "PTV" in the text file.
                                        if (excelStructName.ToUpper().Contains("PTV"))
                                        {
                                            ptvVolume = target.Volume;
                                            break;
                                        }
                                    }

                                    for (int structCount = 0; structCount < numOfStructures; structCount++)
                                    {
                                        string[] structLine = lines[index].Split('|');
                                        string eclipseStructID = structLine[1].Trim().ToUpper();
                                        string excelStructName = structLine[0].Trim();

                                        string[] structQueries = structLine[2].Trim().Split(new char[] { ' ', '\t' },
                                                                                            StringSplitOptions.RemoveEmptyEntries);

                                        Structure target = EclipseExtensions.GetStructureByID(planSum, eclipseStructID);

                                        if (target == null)
                                        {
                                            errorList.Add(new Tuple<string, string>(planSum.Id, eclipseStructID));
                                        }

                                        // We store the structure's computed DVH so we don't need
                                        // to recompute it each time. We store the DVHData in the
                                        // for-loop and not in a dictionary because no structure
                                        // should appear twice in our query.
                                        DVHData structDvh = null;

                                        // Decreasing the step size results in more points in the DVH curve data, and although it is
                                        // more accurate, it also increases the computing time. 
                                        // The max, min, and mean doses are unaffected by the step size.
                                        // The default step size Varian uses in calculating the DVH is 0.001.
                                        double stepSize = 0.01;

                                        // Calculate the DVHData for the structure.
                                        if (IsNullOrEmpty(structQueries))
                                        {
                                            // If we only need the max, min and mean doses, we calculate an inaccurate DVH.
                                            structDvh = planSum.GetDVHCumulativeData(target, DoseValuePresentation.Absolute,
                                                VolumePresentation.AbsoluteCm3, 0.1);
                                        }
                                        else
                                        {
                                            // If we need to calculate other doses, we calculate an accurate DVH.
                                            structDvh = planSum.GetDVHCumulativeData(target, DoseValuePresentation.Absolute,
                                                VolumePresentation.AbsoluteCm3, stepSize);
                                        }

                                        // Every structure should by default output the max, min, and mean doses.
                                        queriesByPlan.AddDVHQuery(new DVHQuery("svol", planSum, target, totalPrescribedDose, excelStructName, ptvVolume, structDvh));
                                        queriesByPlan.AddDVHQuery(new DVHQuery("dmax", planSum, target, totalPrescribedDose, excelStructName, ptvVolume, structDvh));
                                        queriesByPlan.AddDVHQuery(new DVHQuery("dmin", planSum, target, totalPrescribedDose, excelStructName, ptvVolume, structDvh));
                                        queriesByPlan.AddDVHQuery(new DVHQuery("dmean", planSum, target, totalPrescribedDose, excelStructName, ptvVolume, structDvh));

                                        // Output the equivalent sphere diameter for specific target volumes.
                                        if (excelStructName.ToUpper().Contains("CTV") || eclipseStructID.Contains("CTV") ||
                                            excelStructName.ToUpper().Contains("GTV") || eclipseStructID.Contains("GTV") ||
                                            excelStructName.ToUpper().Contains("PTV") || eclipseStructID.Contains("PTV"))
                                        {
                                            queriesByPlan.AddDVHQuery(new DVHQuery("eqsd", planSum, target, totalPrescribedDose, excelStructName, ptvVolume, structDvh));
                                        }

                                        foreach (string query in structQueries)
                                        {
                                            DVHQuery currentQuery = new DVHQuery(query, planSum, target, totalPrescribedDose, excelStructName, ptvVolume, structDvh);
                                            queriesByPlan.AddDVHQuery(currentQuery);
                                        }

                                        index++; // lines[index] == Excel Structure Name : Eclipse Structure ID : Queries || Course
                                    }
                                }
                                else
                                {
                                    //Self-exiting messagebox pop up (change timeout timer at top to change exit time)
                                    DialogResult planDialog = MessageBoxEx.Show($"Plan {eclipsePlanID} could not be found. Would you like to continue?", "Warning",
                                                                              MessageBoxButtons.OKCancel, MessageBoxIcon.Warning, timeout);
                                    //write to failed extractions with ID and reason
                                    using (System.IO.StreamWriter failedExtractionFile =
                                        new System.IO.StreamWriter(failedExtractionLoc, true))
                                    {
                                        failedExtractionFile.Write(usableID);
                                        failedExtractionFile.WriteLine($" -- Plan, {eclipsePlanID}, could not be found.");
                                    }

                                    if (planDialog == DialogResult.Cancel)
                                    {
                                        //Self-exiting messagebox pop up (change timeout timer at top to change exit time)
                                        MessageBoxEx.Show($"Dose extraction could not be completed for {patient.LastName}, {patient.FirstName} ({patient.Id}).", "Extraction Failed",
                                                        MessageBoxButtons.OK, MessageBoxIcon.Error, timeout);

                                        //write to failed extractions with ID and reason
                                        using (System.IO.StreamWriter failedExtractionFile =
                                            new System.IO.StreamWriter(failedExtractionLoc, true))
                                        {
                                            failedExtractionFile.Write(usableID);
                                            failedExtractionFile.WriteLine($" -- Cancelled after Plan, {eclipsePlanID}, could not be found.");
                                        }
                                        //do not exit code, go to next loop iteration
                                        goto failedException;
                                    }
                                    // Increments the index so that on the next loop we are immediately at the next Course.
                                    // (Skips looking at the structures).
                                    index = index + Convert.ToInt32(lines[index]) + 1;
                                    continue;
                                }
                            }
                        }

                        if (queryList.Count == 0)
                        {
                            //Self-exiting messagebox pop up (change timeout timer at top to change exit time)
                            MessageBoxEx.Show($"Dose extraction could not be completed for {patient.LastName}, {patient.FirstName} ({patient.Id}).\n" +
                                        $"None of the queried courses/plans/structures could be found.", "Extraction Failed",
                                        MessageBoxButtons.OK, MessageBoxIcon.Error, timeout);

                            //write to failed extractions with ID and reason
                            using (System.IO.StreamWriter failedExtractionFile =
                                new System.IO.StreamWriter(failedExtractionLoc, true))
                            {
                                failedExtractionFile.Write(usableID);
                                failedExtractionFile.WriteLine(" -- None of the queried courses/plans/structures for the patient could be found.");
                            }
                            //do not exit code, go to next loop iteration
                            goto failedException;
                        }

                        if (errorList.Count > 0)
                        {
                            StringBuilder sb = new StringBuilder();
                            string prev = "";

                            foreach (Tuple<string, string> error in errorList)
                            {
                                if (error.Item1 != prev)
                                {
                                    sb.AppendLine("");
                                    sb.AppendLine(error.Item1 + ":");
                                }
                                sb.AppendLine('\t' + error.Item2);

                                prev = error.Item1;
                            }

                            sb.AppendLine("");

                            //Self-exiting messagebox pop up (change timeout timer at top to change exit time)
                            DialogResult errorDialog = MessageBoxEx.Show($"The following structures could not be found:\n{sb.ToString()}Would you like to continue?", "Warning",
                                                                       MessageBoxButtons.OKCancel, MessageBoxIcon.Warning, 5000);
                            //write to failed extractions with ID and reason
                            using (System.IO.StreamWriter failedExtractionFile =
                                new System.IO.StreamWriter(failedExtractionLoc, true))
                            {
                                failedExtractionFile.Write(usableID);
                                failedExtractionFile.WriteLine($" -- Structures, {sb.ToString()}, could not be found.");
                            }

                            if (errorDialog == DialogResult.Cancel)
                            {
                                //Self-exiting messagebox pop up (change timeout timer at top to change exit time)
                                MessageBoxEx.Show($"Dose extraction could not be completed for {patient.LastName}, {patient.FirstName} ({patient.Id}).", "Extraction Failed",
                                                        MessageBoxButtons.OK, MessageBoxIcon.Error, timeout);

                                //write to failed extractions with ID and reason
                                using (System.IO.StreamWriter failedExtractionFile =
                                    new System.IO.StreamWriter(failedExtractionLoc, true))
                                {
                                    failedExtractionFile.Write(usableID);
                                    failedExtractionFile.WriteLine($" -- Cancelled after structures, {sb.ToString()}, could not be found.");
                                }
                                //do not exit code, go to next loop iteration
                                goto failedException;
                            }
                        }
						Console.Write($"\rCompleted Dose Extraction for {usableID}. Loading next patient...  ");
						//string idListLoc = inputInfo.Substring(0, 71);
						//string idList = $@"{idListLoc}\{structure} Pending Extractions List.txt";
						//
						//idList.Remove(count, 8);

						excelHelper.ExportDoseAsExcel(queryList);

                        //Self-exiting messagebox pop up (change timeout timer at top to change exit time)
                        MessageBoxEx.Show($"DVH queries successfully extracted for {patient.LastName}, {patient.FirstName} ({patient.Id}).", "Extraction Complete", timeout);
                        //failed exception loop identify
                        failedException:
                        //close the patient so that we can open the next one on the next loop
                        app.ClosePatient();
                    }
                    else
                    {
                        //Self-exiting messagebox pop up (change timeout timer at top to change exit time)
                        MessageBoxEx.Show($"Template does not exist for {patient.Id}.", "Extraction Failed",
                                        MessageBoxButtons.OK, MessageBoxIcon.Error, timeout);
                        //if extraction fails writes to failed extractions w/ ID and reason
                        using (System.IO.StreamWriter failedExtractionFile =
                            new System.IO.StreamWriter(failedExtractionLoc, true))
                        {
                            failedExtractionFile.Write(usableID);
                            failedExtractionFile.WriteLine(" -- Template did not exist in Patient Template Location.");
                        }

                        //count = count + 8;
                    }
                }
                //close file so that we don't run into files left open
                file.Close();

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
			}
            else
            {
                MessageBox.Show("Patient Template for have not been created yet.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            File.WriteAllText(workingLocExternal, structure);
        }

        // Returns true if an IEnumerable<T> is not null and not empty.
        public static bool IsNullOrEmpty<T>(IEnumerable<T> data)
        {
            return data == null || !data.Any();
        }
    }
}