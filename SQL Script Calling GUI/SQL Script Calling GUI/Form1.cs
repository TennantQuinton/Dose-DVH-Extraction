using System;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;       //microsoft Excel 14 object in references-> COM tab


namespace SQL_Script_Calling_GUI
{
	public partial class Form1 : Form
	{
		private int originalWidth;
		private int originalHeight;

		public Form1()
		{
			InitializeComponent();
			originalWidth = this.Width;
			originalHeight = this.Height;

            groupBoxAdvQ.Location = new System.Drawing.Point(12, 294);

            dateTimePickerStart.Format = DateTimePickerFormat.Custom;
            dateTimePickerEnd.Format = DateTimePickerFormat.Custom;
            dateTimePickerStart.CustomFormat = "yyyy-MM-dd";
            dateTimePickerEnd.CustomFormat = "yyyy-MM-dd";

            //dateTimePickerStart.Value = System.DateTime.Now;
            //dateTimePickerEnd.Value = System.DateTime.Now;

			this.Width = 333;
			this.Height = 390;

			this.MaximumSize = new System.Drawing.Size(333, 390);

			if (textBoxServer.Text.Length > 0 && textBoxDatabase.Text.Length > 0 && textBoxUserID.Text.Length > 0 && textBoxPassword.Text.Length > 0)
			{
				groupBoxQuery.Enabled = true;
			}
			else
			{
				groupBoxQuery.Enabled = false;
			}

			if (textBoxServer.Text.Length > 0)
			{
				textBoxDatabase.Enabled = true;
				if (textBoxDatabase.Text.Length > 0)
				{
					textBoxUserID.Enabled = true;
					if (textBoxUserID.Text.Length > 0)
					{
						textBoxPassword.Enabled = true;
						if (textBoxPassword.Text.Length > 0)
						{
							dateTimePickerStart.Enabled = true;
							if (dateTimePickerStart.Checked == true)
							{
								dateTimePickerEnd.Enabled = true;
								if (dateTimePickerEnd.Checked == true)
								{
									comboBoxSite.Enabled = true;
									if (comboBoxSite.SelectedIndex > 0)
									{
										textBoxExcel.Enabled = true;
										if (textBoxExcel.Text.Length > 0)
										{
											buttonQuery.Enabled = true;
										}
									}
								}
							}
						}
					}
				}
			}
		}

		private void button1_Click(object sender, EventArgs e)
		{
			string serverName = textBoxServer.Text;
			string databaseName = textBoxDatabase.Text;
			string userID = textBoxUserID.Text;
			string password = textBoxPassword.Text;
			string sqlConnectionString = ($"Server={serverName}; Database={databaseName}; User Id={userID}; Password={password};");
			string databaseQuery = @"\\dc3-pr-files\MedPhysics Backup\Data Extractions\Automation\Database Query.sql";

			// Create the connection to the resource!
			// This is the connection, that is established and
			// will be available throughout this block.

			string startDate = dateTimePickerStart.Text;
			string endDate = dateTimePickerEnd.Text;
			string structure = comboBoxSite.Text;
			string excelName = textBoxExcel.Text;
			string volumeCode = textBoxVolCode.Text;
			string srsComment = "";

			if (structure.IndexOf("brai", StringComparison.OrdinalIgnoreCase) >= 0)
			{
				srsComment = "";
			}
			else if (structure.IndexOf("brea", StringComparison.OrdinalIgnoreCase) >= 0)
			{
				srsComment = "--";
			}
			else if (structure.IndexOf("lung", StringComparison.OrdinalIgnoreCase) >= 0)
			{
				srsComment = "--";
			}
			else if (structure.IndexOf("pros", StringComparison.OrdinalIgnoreCase) >= 0)
			{
				srsComment = "--";
			}
            else if (structure.IndexOf("All Patients", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                srsComment = "--";
            }
            else
			{
				MessageBox.Show("No valid structure selected.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				goto exit;
			}

			string gantryComment = "--";
			string gantryLike = "";
			string gantryCond = "";

			string MLCComment = "--";
			string MLCLike = "";
			string MLCCond = "";

			string doseComment = "--";
			string doseSign = "";
			string doseNumb = "";

			string fractComment = "--";
			string fractSign = "";
			string fractNumb = "";

			string courseIDComment = "--";
			string courseIDLike = "";
			string courseIDName = "";

            string patIDComment = "--";
            string patIDLike = "";
            string patIDCond = "";

			if (checkGantryRot.Checked == true)
			{
				gantryComment = "";
				gantryLike = comboGantLike.Text;
				gantryCond = comboGantryCond.Text;
			}
			if (checkMLCPlan.Checked == true)
			{
				MLCComment = "";
				MLCLike = comboMLCLike.Text;
				MLCCond = comboMLCCon.Text;
			}
			if (checkPresDose.Checked == true)
			{
				doseComment = "";
				doseSign = comboDoseEq.Text;
				doseNumb = numericDose.Value.ToString();
			}
			if (checkNFract.Checked == true)
			{
				fractComment = "";
				fractSign = comboFractEq.Text;
				fractNumb = numericFract.Value.ToString();
			}
            if (checkPatID.Checked == true)
            {
                patIDComment = "";
                patIDLike = comboPatIDLike.Text;
                patIDCond = textPatIDCond.Text;
            }

			if (volumeCode.Length > 0)
			{
				using (SqlConnection conn = new SqlConnection())
				{
					// Create the connectionString
					// Trusted_Connection is used to denote the connection uses Windows Authentication
					conn.ConnectionString = sqlConnectionString; //"Server=grc505n; Database=variansystem; User Id = reports; Password = reports";

					conn.Open();
					// Create the command
					string sqlCommandString = $@"
select distinct
(p.PatientId) PatientId,
(c.CourseId) CourseId,
(ps.PlanSetupId) Plan_Name,
dvh.Structures,
mlcp.MLCPlanType,
--(p.DateOfBirth) Date_Of_Birth,
(vc.VolumeCode) Body_Region,
rtp.PrescribedDose,
rtp.NoFractions,
--f.GantryRtnDirection,
(ps.CreationDate) Plan_CreationDate

from PlanSetup ps
join Course c on c.CourseSer = ps.CourseSer
inner join Patient p on p.PatientSer = c.PatientSer
inner join RTPlan rtp on rtp.PlanSetupSer = ps.PlanSetupSer
inner join DoseContribution dc
	on rtp.RTPlanSer = dc.RTPlanSer
inner join RefPoint rp
	on dc.RefPointSer = rp.RefPointSer
inner join PatientVolume pv
	on rp.PatientVolumeSer = pv.PatientVolumeSer
inner join VolumeCode vc
	on pv.VolumeCodeSer = vc.VolumeCodeSer
inner join Radiation r
	on ps.PlanSetupSer = r.PlanSetupSer
inner join ExternalFieldCommon fc
	on r.RadiationSer = fc.RadiationSer
inner join MLCPlan mlcp
	on r.RadiationSer = mlcp.RadiationSer
inner join DVH dvh
	on rtp.PlanSetupSer = dvh.PlanSetupSer

where 1=1
and cast(ps.CreationDate as date) between '{startDate}' and '{endDate}'
and vc.VolumeCode like ('{volumeCode}')
{srsComment}and ps.PlanSetupId like '%SRS%'
and p.PatientId not like '%$%'
{patIDComment}and p.PatientId {patIDLike} '{patIDCond}'
and c.CourseId not like ('%QA%')
{gantryComment}and f.GantryRtnDirection {gantryLike} '{gantryCond}'
{MLCComment}and mlcp.MLCPlanType {MLCLike} '{MLCCond}'
and c.CourseId in ('1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12', '13', '14', '15', '16', '17', '18', '19', '20')
{doseComment}and rtp.PrescribedDose {doseSign} '{doseNumb}'
{fractComment}and rtp.NoFractions {fractSign} '{fractNumb}'
order by p.PatientId";

					SqlCommand command = new SqlCommand(sqlCommandString, conn);
					// Add the parameters.
					command.Parameters.Add(new SqlParameter("0", 1));

					string sqlScriptFile = $@"\\dc3-pr-files\MedPhysics Backup\Data Extractions\Automation\Database Query.sql";
					File.WriteAllText(sqlScriptFile, sqlCommandString);

					Microsoft.Office.Interop.Excel.Application oXL;
					Microsoft.Office.Interop.Excel._Workbook oWB;
					Microsoft.Office.Interop.Excel._Worksheet oSheet;
					Microsoft.Office.Interop.Excel.Range oRng;

					object misvalue = System.Reflection.Missing.Value;
					//Start Excel and get Application object.
					oXL = new Microsoft.Office.Interop.Excel.Application();
					oXL.DisplayAlerts = false;
					oXL.Visible = true;

					//Get a new workbook.
					oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Add(""));
					oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;

					using (SqlDataReader reader = command.ExecuteReader())
					{
						//Add table headers going cell by cell.
						oSheet.Cells[1, 1] = "PatientId";
						oSheet.Cells[1, 2] = "CourseId";
						oSheet.Cells[1, 3] = "Plan_Name";
						oSheet.Cells[1, 4] = "Structures";
						oSheet.Cells[1, 5] = "MLCPlanType";
						oSheet.Cells[1, 6] = "Body_Region";
						oSheet.Cells[1, 7] = "PrescribedDose";
						oSheet.Cells[1, 8] = "No. Fractions";
						oSheet.Cells[1, 9] = "Plan_CreationDate";
                        //oSheet.Cells[1, 10] = "GantryRtnDirection";

                        Excel.Range final = oSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                        int colCountMax = final.Column;

						int rowCount = 2;
						while (reader.Read())
						{
							for (int colCount = 1; colCount <= colCountMax; colCount++)
							{
								oSheet.Cells[rowCount, colCount] = $"{reader[colCount - 1]}";
							}
							rowCount++;
						}
					}

					oXL.Visible = true;
					oXL.UserControl = true;

                    Excel.Range xlRangePat = oSheet.UsedRange;
                    int rowCountPat = xlRangePat.Rows.Count;

                    Excel.Range last = oSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);

                    oSheet.Range[$"A1:H{last.Row}"].NumberFormat = "@";
                    oSheet.Range[$"I2:I{last.Row}"].NumberFormat = "yyyy-mm-dd hh:mm";

                    string excelLoc = $@"\\dc3-pr-files\MedPhysics Backup\Data Extractions\Automation\SQL Database Spreadsheets\{structure}\{excelName}.xlsx";
                    oWB.SaveAs(excelLoc, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                    false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

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
                            else if (cell.Length == 3)
                            {
                                string cellNew = $"00000{cell}";
                                xlRangePat.Cells[i, 1].Value = cellNew;
                            }
                            else if (cell.Length == 2)
                            {
                                string cellNew = $"000000{cell}";
                                xlRangePat.Cells[i, 1].Value = cellNew;
                            }
                            else if (cell.Length == 1)
                            {
                                string cellNew = $"0000000{cell}";
                                xlRangePat.Cells[i, 1].Value = cellNew;
                            }
                        }
                    }

					oWB.SaveAs(excelLoc, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
						false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
						Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                    MessageBox.Show($"Database successfully queried for {structure} from {startDate} until {endDate}", "Query Complete", MessageBoxButtons.OK);

					//oWB.Close();
					//oXL.Quit();


				}
			}
			else
			{
				MessageBox.Show("No volume code input.", "Volume Code", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		exit:;
		}

		private void textBoxServer_TextChanged(object sender, EventArgs e)
		{

			if (textBoxServer.Text.Length > 0)
			{
				textBoxDatabase.Enabled = true;
			}
			else
			{
				textBoxDatabase.Enabled = false;
				textBoxDatabase.Text = "";
			}
		}

		private void textBoxDatabase_TextChanged(object sender, EventArgs e)
		{

			if (textBoxDatabase.Text.Length > 0)
			{
				textBoxUserID.Enabled = true;
			}
			else
			{
				textBoxUserID.Enabled = false;
				textBoxUserID.Text = "";
			}
		}

		private void textBoxUserID_TextChanged(object sender, EventArgs e)
		{

			if (textBoxUserID.Text.Length > 0)
			{
				textBoxPassword.Enabled = true;
			}
			else
			{
				textBoxPassword.Enabled = false;
				textBoxPassword.Text = "";
			}
		}

		private void textBoxPassword_TextChanged(object sender, EventArgs e)
		{
			if (textBoxServer.Text.Length > 0 && textBoxDatabase.Text.Length > 0 && textBoxUserID.Text.Length > 0 && textBoxPassword.Text.Length > 0)
			{
				groupBoxQuery.Enabled = true;
			}
			else
			{
				groupBoxQuery.Enabled = false;
			}

			if (textBoxPassword.Text.Length > 0)
			{
				dateTimePickerStart.Enabled = true;
			}
			else
			{
				dateTimePickerStart.Enabled = false;
				dateTimePickerStart.Text = "";
			}
		}

		private void groupBoxDatabase_Enter(object sender, EventArgs e)
		{

		}

		private void textBoxStart_TextChanged(object sender, EventArgs e)
		{

		}

		private void textBoxEnd_TextChanged(object sender, EventArgs e)
		{

		}

		private void textBoxStructure_TextChanged(object sender, EventArgs e)
		{

		}

		private void textBoxExcel_TextChanged(object sender, EventArgs e)
		{
			if (textBoxExcel.Text.Length > 0 && textBoxVolCode.Text.Length > 0)
			{
				buttonQuery.Enabled = true;
			}
			else
			{
				buttonQuery.Enabled = false;
			}
		}

		private void groupBoxQuery_Enter(object sender, EventArgs e)
		{
			
		}

		private void dateTimePickerStart_ValueChanged(object sender, EventArgs e)
		{
			if (dateTimePickerStart.Text.Length > 0)
			{
				dateTimePickerEnd.Enabled = true;
			}
			else
			{
				dateTimePickerEnd.Enabled = false;
				dateTimePickerEnd.Text = "";
			}
		}

		private void dateTimePickerEnd_ValueChanged(object sender, EventArgs e)
		{
			if (dateTimePickerEnd.Text.Length > 0)
			{
				comboBoxSite.Enabled = true;
				comboBoxSite.SelectedIndex = 0;
			}
			else
			{
				comboBoxSite.Enabled = false;
				comboBoxSite.SelectedIndex = 0;
			}
		}

		private void comboBoxStructure_SelectedIndexChanged(object sender, EventArgs e)
		{
			textBoxVolCode.Text = "";
			checkGantryRot.Checked = false;
			checkMLCPlan.Checked = false;
			checkPresDose.Checked = false;
			checkNFract.Checked = false;

			string startDate = dateTimePickerStart.Text;
			string endDate = dateTimePickerEnd.Text;

			string startYear = startDate.Substring((startDate.Length - 4), 4);
			string endYear = endDate.Substring((endDate.Length - 4), 4);

            if (startYear.Contains("-"))
            {
                startYear = startDate.Substring(0, 4);
            }

            if (endYear.Contains("-"))
            {
                endYear = endDate.Substring(0, 4);
            }

            if (comboBoxSite.Text == "Brain")
			{
				textBoxExcel.Text = $"SRS_{startYear}-{endYear}";
			}
			else if (comboBoxSite.Text == "Breast")
			{
				textBoxExcel.Text = $"Breast_{startYear}-{endYear}";
			}
			else if (comboBoxSite.Text == "Lung")
			{
				textBoxExcel.Text = $"Lung_{startYear}-{endYear}";
			}
			else if (comboBoxSite.Text == "Prostate")
			{
				textBoxExcel.Text = $"Pros_{startYear}-{endYear}";
			}
            else if (comboBoxSite.Text == "All Patients")
            {
                textBoxExcel.Text = $"ALL_{startYear}-{endYear}";
                textBoxVolCode.Text = "%";
            }
			else
			{
				textBoxExcel.Text = $"";
			}

			if (comboBoxSite.SelectedIndex != 0)
			{
				textBoxVolCode.Enabled = true;
			}
			else
			{
				textBoxVolCode.Enabled = false;
				textBoxVolCode.Text = "";
			}

			
		}

		private void Form1_Load(object sender, EventArgs e)
		{

		}

		private void button1_Click_1(object sender, EventArgs e)
		{
			string exePath = @"\\dc3-pr-files\MedPhysics Backup\Coop Students\2019\Term 3 - Autumn\Quinton Tennant\Release Copy\GUI Test\GUI Test\bin\Debug\GUI Test.exe";

			Process process = new Process();
			// Pass your exe file path here.
			string path = exePath;
			string fileName = Path.GetFileName(path);
			// Get the precess that already running as per the exe file name.
			Process[] processName = Process.GetProcessesByName(fileName.Substring(0, fileName.LastIndexOf('.')));
			if (processName.Length > 0)
			{
				MessageBox.Show("Instance of Extract Dose/DVH already open", "exe Already Open", MessageBoxButtons.OK, MessageBoxIcon.Warning);
			}
			else
			{
				string extractExe = @"\\dc3-pr-files\MedPhysics Backup\Data Extractions\Automation\Executables\Extraction Automation GUI.lnk";
				System.Diagnostics.Process.Start(extractExe);
			}
		}

		private void textBoxVolCode_TextChanged(object sender, EventArgs e)
		{
			if (textBoxVolCode.Text.Length > 0)
			{
				textBoxExcel.Enabled = true;
			}
			else
			{
				textBoxExcel.Enabled = false;
			}

			if (textBoxExcel.Text.Length > 0 && textBoxVolCode.Text.Length > 0)
			{
				buttonQuery.Enabled = true;
			}
			else
			{
				buttonQuery.Enabled = false;
			}
		}

		private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
		{
			labelExcel.Location = new System.Drawing.Point(9, 451);
			textBoxExcel.Location = new System.Drawing.Point(82, 448);

			groupBoxAdvQ.Visible = true;
			groupBoxAdvQ.Enabled = true;
			buttonQuery.Location = new System.Drawing.Point(179, 470);
			linkLabelAdvQ.Visible = false;
			linkLabelAdvQ.Enabled = false;
			linkLabelCancel.Visible = true;
			linkLabelCancel.Enabled = true;
			this.MinimumSize = new System.Drawing.Size(333, 550);
			this.MaximumSize = new System.Drawing.Size(333, 550);

			this.Width = 333;
			this.Height = 550;

			checkGantryRot.Checked = false;
			checkMLCPlan.Checked = false;
			checkPresDose.Checked = false;
			checkNFract.Checked = false;
		}

		private void linkLabelBaseQ_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
		{
			labelExcel.Location = new System.Drawing.Point(9, 294);
			textBoxExcel.Location = new System.Drawing.Point(82, 291);

			groupBoxAdvQ.Visible = false;
			groupBoxAdvQ.Enabled = false;
			buttonQuery.Location = new System.Drawing.Point(179, 315);
			linkLabelAdvQ.Visible = true;
			linkLabelAdvQ.Enabled = true;
			linkLabelCancel.Visible = false;
			linkLabelCancel.Enabled = false;
			this.MinimumSize = new System.Drawing.Size(333, 390);
			this.MaximumSize = new System.Drawing.Size(333, 390);

			this.Width = 333;
			this.Height = 390;

			checkGantryRot.Checked = false;
			checkMLCPlan.Checked = false;
			checkPresDose.Checked = false;
			checkNFract.Checked = false;

			comboGantLike.SelectedIndex = 0;
			comboMLCLike.Text = "";
			comboDoseEq.Text = "";
			comboFractEq.Text = "";

			comboGantryCond.Text = "";
			comboMLCCon.Text = "";
			numericDose.Value = 0;
			numericFract.Value = 0;

		}

		private void numericUpDown1_ValueChanged(object sender, EventArgs e)
		{

		}

		private void numericUpDown2_ValueChanged(object sender, EventArgs e)
		{

		}

		private void labelExcel_Click(object sender, EventArgs e)
		{

		}

		private void label2_Click(object sender, EventArgs e)
		{

		}

		private void textBox1_TextChanged(object sender, EventArgs e)
		{

		}

		private void checkGantryRot_CheckedChanged(object sender, EventArgs e)
		{
			int excelLength = textBoxExcel.Text.Length;
			if (excelLength > 0 && !textBoxExcel.Text.Contains("GR") && checkGantryRot.Checked)
			{
				textBoxExcel.Text = textBoxExcel.Text.Insert(excelLength, ", GR");
			}

			if (!checkGantryRot.Checked)
			{
				int GRIndex = textBoxExcel.Text.IndexOf(", GR");
				if (GRIndex != -1)
				{
					textBoxExcel.Text = textBoxExcel.Text.Remove(GRIndex, 4);
				}
			}
		}

		private void checkMLCPlan_CheckedChanged(object sender, EventArgs e)
		{
			int excelLength = textBoxExcel.Text.Length;
			if (excelLength > 0 && !textBoxExcel.Text.Contains("MPT") && checkMLCPlan.Checked)
			{
				textBoxExcel.Text = textBoxExcel.Text.Insert(excelLength, ", MPT");
			}

			if (!checkMLCPlan.Checked)
			{
				int MPTIndex = textBoxExcel.Text.IndexOf(", MPT");
				if (MPTIndex != -1)
				{
					textBoxExcel.Text = textBoxExcel.Text.Remove(MPTIndex, 5);
				}
			}
		}

		private void checkPresDose_CheckedChanged(object sender, EventArgs e)
		{
			int excelLength = textBoxExcel.Text.Length;
			if (excelLength > 0 && !textBoxExcel.Text.Contains("PD") && checkPresDose.Checked)
			{
				textBoxExcel.Text = textBoxExcel.Text.Insert(excelLength, ", PD");
			}

			if (!checkPresDose.Checked)
			{
				int PDIndex = textBoxExcel.Text.IndexOf(", PD");
				if (PDIndex != -1)
				{
					textBoxExcel.Text = textBoxExcel.Text.Remove(PDIndex, 4);
				}
			}
		}

		private void checkNFract_CheckedChanged(object sender, EventArgs e)
		{
			int excelLength = textBoxExcel.Text.Length;
			if (excelLength > 0 && !textBoxExcel.Text.Contains("NoF") && checkNFract.Checked)
			{
				textBoxExcel.Text = textBoxExcel.Text.Insert(excelLength, ", NoF");
			}

			if (!checkNFract.Checked)
			{
				int NoFIndex = textBoxExcel.Text.IndexOf(", NoF");
				if (NoFIndex != -1)
				{
					textBoxExcel.Text = textBoxExcel.Text.Remove(NoFIndex, 5);
				}
			}
		}

        private void checkPatID_CheckedChanged(object sender, EventArgs e)
        {
            int excelLength = textBoxExcel.Text.Length;
            if (excelLength > 0 && !textBoxExcel.Text.Contains("PID") && checkPatID.Checked)
            {
                textBoxExcel.Text = textBoxExcel.Text.Insert(excelLength, ", PID");
            }

            if (!checkPatID.Checked)
            {
                int PIDIndex = textBoxExcel.Text.IndexOf(", PID");
                if (PIDIndex != -1)
                {
                    textBoxExcel.Text = textBoxExcel.Text.Remove(PIDIndex, 5);
                }
            }
        }
    }
}
