using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO;

namespace GUI_Test
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

			singleTemplateButton.Location = new System.Drawing.Point(20, 19);
        }
        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex != 0)
            {
                groupBoxStructure.Enabled = true;
                checkBoxPros.Enabled = true;
				checkBoxLung.Enabled = true;
				checkBoxBrea.Enabled = true;
				checkBoxBrai.Enabled = true;
				checkBoxInput.Enabled = true;
                textBoxInput.Enabled = true;

				checkBoxPros.Checked = false;
				checkBoxLung.Checked = false;
				checkBoxBrea.Checked = false;
				checkBoxBrai.Checked = false;
				checkBoxInput.Checked = false;
				linkLabelDesAll.Visible = false;
				linkLabelSelAll.Visible = true;
				linkLabelDesAll.Enabled = false;
				linkLabelSelAll.Enabled = true;

                archiveButton.Enabled = false;
                failedButton.Enabled = false;
                spreadsheetButton.Enabled = false;

				checkBoxDose.Enabled = true;
				checkBoxDose.Checked = false;
				checkBoxDVH.Enabled = true;
				checkBoxDVH.Checked = false;
            }
            else
            {
                groupBoxStructure.Enabled = false;
				checkBoxPros.Enabled = false;
				checkBoxLung.Enabled = false;
                checkBoxBrea.Enabled = false;
                checkBoxBrai.Enabled = false;
                checkBoxInput.Enabled = false;
                textBoxInput.Enabled = false;
                textBoxID.Enabled = false;
                labelID.Visible = true;
                textBoxID.Text = "";
                labelID.Enabled = false;

                checkBoxPros.Checked = false;
                checkBoxLung.Checked = false;
                checkBoxBrea.Checked = false;
                checkBoxBrai.Checked = false;
                checkBoxInput.Checked = false;
				linkLabelDesAll.Visible = false;
				linkLabelSelAll.Visible = true;
				linkLabelDesAll.Enabled = false;
				linkLabelSelAll.Enabled = true;

				userInputGroup.Enabled = false;
                editingGroup.Enabled = false;
                runExeButton.Enabled = false;
                failedButton.Enabled = false;
                archiveButton.Enabled = false;
                spreadsheetButton.Enabled = false;

				checkPlanSumBrai.Checked = false;
				checkPlanSumBrea.Checked = false;
				checkPlanSumInput.Checked = false;
				checkPlanSumLung.Checked = false;
				checkPlanSumPros.Checked = false;


				textBoxInput.Text = "";

				checkBoxDose.Enabled = false;
				checkBoxDVH.Enabled = false;
				checkBoxDose.Checked = false;
				checkBoxDVH.Checked = false;
            }

            if (comboBox1.Text == "Clear Templates Archive")
            {
                groupBoxStructure.Enabled = false;
                checkBoxPros.Enabled = false;
                checkBoxLung.Enabled = false;
                checkBoxBrea.Enabled = false;
                checkBoxBrai.Enabled = false;
                checkBoxInput.Enabled = false;
                textBoxInput.Enabled = false;

                checkBoxPros.Checked = false;
                checkBoxLung.Checked = false;
                checkBoxBrea.Checked = false;
                checkBoxBrai.Checked = false;
                checkBoxInput.Checked = false;
				linkLabelDesAll.Visible = false;
				linkLabelSelAll.Visible = true;
				linkLabelDesAll.Enabled = false;
				linkLabelSelAll.Enabled = true;

				userInputGroup.Enabled = false;
                editingGroup.Enabled = false;
                runExeButton.Enabled = true;
                failedButton.Enabled = false;
                archiveButton.Enabled = true;
                spreadsheetButton.Enabled = false;

				checkBoxDose.Enabled = true;
				checkBoxDVH.Enabled = true;
				checkBoxDose.Checked = false;
				checkBoxDVH.Checked = false;
			}

            if (comboBox1.Text == "Extraction (Single Patient)")
            {
                textBoxID.Enabled = true;
                groupBoxStructure.Enabled = false;
                checkBoxPros.Checked = false;
                checkBoxBrea.Checked = false;
                checkBoxLung.Checked = false;
                checkBoxBrai.Checked = false;
                checkBoxInput.Checked = false;
				linkLabelDesAll.Visible = false;
				linkLabelSelAll.Visible = true;
				linkLabelDesAll.Enabled = false;
				linkLabelSelAll.Enabled = true;

				userInputGroup.Enabled = false;
                editingGroup.Enabled = false;

				groupBoxPlanSum.Enabled = false;
				groupBoxPlanSum.Visible = false;
            }
            else
            {
                textBoxID.Text = "";
                textBoxID.Enabled = false;
                labelID.Enabled = false;
                singleTemplateButton.Visible = false;
                singleTemplateButton.Enabled = false;
            }

			if (comboBox1.Text == "Patient Template Creation" || comboBox1.Text == "Run Template and Extraction")
			{
				groupBoxPlanSum.Enabled = true;
				groupBoxPlanSum.Visible = true;
			}
			else if (comboBox1.Text == "Extraction" || comboBox1.Text == "Extraction (Single Patient)")
			{
				groupBoxPlanSum.Enabled = false;
				groupBoxPlanSum.Visible = false;
			}
        }

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        private void textBox1_TextChanged_1(object sender, EventArgs e)
        {

        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            string READMELoc = @"\\dc3-pr-files\MedPhysics Backup\Data Extractions\Automation\README.pdf";
            System.Diagnostics.Process.Start(READMELoc);
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

		private void checkBoxBrai_CheckedChanged(object sender, EventArgs e)
		{
			if (comboBox1.SelectedIndex != 0)
			{
				userInputGroup.Enabled = true;
				runExeButton.Enabled = true;

				if (comboBox1.Text == "Extraction" || comboBox1.Text == "Extraction (Single Patient)")
				{
					editingGroup.Enabled = true;
					groupBoxPlanSum.Enabled = false;
					groupBoxPlanSum.Visible = false;
					checkPlanSumBrai.Checked = false;
					checkPlanSumBrea.Checked = false;
					checkPlanSumInput.Checked = false;
					checkPlanSumLung.Checked = false;
					checkPlanSumPros.Checked = false;
				}

				if (comboBox1.Text == "Patient Template Creation" || comboBox1.Text == "Run Template and Extraction")
				{
					editingGroup.Enabled = false;
					groupBoxPlanSum.Enabled = true;
					groupBoxPlanSum.Visible = true;
					checkPlanSumPros.Enabled = true;
					checkPlanSumInput.Enabled = true;
					//checkPlanSumBrai.Enabled = true;
					//checkPlanSumBrea.Enabled = true;
                    //checkPlanSumLung.Enabled = true;
				}

				else
				{
					groupBoxPlanSum.Enabled = false;
					groupBoxPlanSum.Visible = false;
					checkPlanSumInput.Enabled = false;
					checkPlanSumPros.Enabled = false;
                    //checkPlanSumBrai.Enabled = true;
                    //checkPlanSumBrea.Enabled = true;
                    //checkPlanSumLung.Enabled = true;

                    checkPlanSumPros.Checked = false;
					checkPlanSumInput.Checked = false;
                    checkPlanSumBrai.Checked = false;
                    checkPlanSumBrea.Checked = false;
                    checkPlanSumLung.Checked = false;
                }
			}
			textBoxInput.Text = "";
		}

		private void checkBoxBrea_CheckedChanged(object sender, EventArgs e)
		{
			if (comboBox1.SelectedIndex != 0)
			{
				userInputGroup.Enabled = true;
				runExeButton.Enabled = true;

				if (comboBox1.Text == "Extraction" || comboBox1.Text == "Extraction (Single Patient)")
				{
					editingGroup.Enabled = true;
					groupBoxPlanSum.Enabled = false;
					groupBoxPlanSum.Visible = false;
					checkPlanSumBrai.Checked = false;
					checkPlanSumBrea.Checked = false;
					checkPlanSumInput.Checked = false;
					checkPlanSumLung.Checked = false;
					checkPlanSumPros.Checked = false;
				}

				if (comboBox1.Text == "Patient Template Creation" || comboBox1.Text == "Run Template and Extraction")
				{
					editingGroup.Enabled = false;
					groupBoxPlanSum.Enabled = true;
					groupBoxPlanSum.Visible = true;
					checkPlanSumPros.Enabled = true;
					checkPlanSumInput.Enabled = true;
					checkPlanSumPros.Enabled = true;
					checkPlanSumInput.Enabled = true;
				}

				else
				{
					groupBoxPlanSum.Enabled = false;
					groupBoxPlanSum.Visible = false;
					checkPlanSumInput.Enabled = false;
					checkPlanSumPros.Enabled = false;
					checkPlanSumPros.Checked = false;
					checkPlanSumInput.Checked = false;
				}
			}
			textBoxInput.Text = "";
		}

		private void checkBoxLung_CheckedChanged(object sender, EventArgs e)
		{
			if (comboBox1.SelectedIndex != 0)
			{
				userInputGroup.Enabled = true;
				runExeButton.Enabled = true;

				if (comboBox1.Text == "Extraction" || comboBox1.Text == "Extraction (Single Patient)")
				{
					editingGroup.Enabled = true;
					groupBoxPlanSum.Enabled = false;
					groupBoxPlanSum.Visible = false;
					checkPlanSumBrai.Checked = false;
					checkPlanSumBrea.Checked = false;
					checkPlanSumInput.Checked = false;
					checkPlanSumLung.Checked = false;
					checkPlanSumPros.Checked = false;
				}

				if (comboBox1.Text == "Patient Template Creation" || comboBox1.Text == "Run Template and Extraction")
				{
					editingGroup.Enabled = false;
					groupBoxPlanSum.Enabled = true;
					groupBoxPlanSum.Visible = true;
					checkPlanSumPros.Enabled = true;
					checkPlanSumInput.Enabled = true;
					checkPlanSumPros.Enabled = true;
					checkPlanSumInput.Enabled = true;
				}

				else
				{
					groupBoxPlanSum.Enabled = false;
					groupBoxPlanSum.Visible = false;
					checkPlanSumInput.Enabled = false;
					checkPlanSumPros.Enabled = false;
					checkPlanSumPros.Checked = false;
					checkPlanSumInput.Checked = false;
				}
			}
			textBoxInput.Text = "";
		}

		private void checkBoxPros_CheckedChanged(object sender, EventArgs e)
		{
			if (comboBox1.SelectedIndex != 0)
			{
				userInputGroup.Enabled = true;
				runExeButton.Enabled = true;

				if (comboBox1.Text == "Extraction" || comboBox1.Text == "Extraction (Single Patient)")
				{
					editingGroup.Enabled = true;
					groupBoxPlanSum.Enabled = false;
					groupBoxPlanSum.Visible = false;
					checkPlanSumBrai.Checked = false;
					checkPlanSumBrea.Checked = false;
					checkPlanSumInput.Checked = false;
					checkPlanSumLung.Checked = false;
					checkPlanSumPros.Checked = false;
				}

				if (comboBox1.Text == "Patient Template Creation" || comboBox1.Text == "Run Template and Extraction")
				{
					editingGroup.Enabled = false;
					groupBoxPlanSum.Enabled = true;
					groupBoxPlanSum.Visible = true;
					checkPlanSumPros.Enabled = true;
					checkPlanSumInput.Enabled = true;
					checkPlanSumPros.Enabled = true;
					checkPlanSumInput.Enabled = true;
				}

				else
				{
					groupBoxPlanSum.Enabled = false;
					groupBoxPlanSum.Visible = false;
					checkPlanSumInput.Enabled = false;
					checkPlanSumPros.Enabled = false;
					checkPlanSumPros.Checked = false;
					checkPlanSumInput.Checked = false;
				}
			}
			textBoxInput.Text = "";
		}

		private void checkBoxInput_CheckedChanged(object sender, EventArgs e)
		{
			if (comboBox1.SelectedIndex != 0)
			{
				userInputGroup.Enabled = true;
				runExeButton.Enabled = true;

				if (comboBox1.Text == "Extraction" || comboBox1.Text == "Extraction (Single Patient)")
				{
					editingGroup.Enabled = true;
					groupBoxPlanSum.Enabled = false;
					groupBoxPlanSum.Visible = false;
					checkPlanSumBrai.Checked = false;
					checkPlanSumBrea.Checked = false;
					checkPlanSumInput.Checked = false;
					checkPlanSumLung.Checked = false;
					checkPlanSumPros.Checked = false;
				}

				if (comboBox1.Text == "Patient Template Creation" || comboBox1.Text == "Run Template and Extraction")
				{
					editingGroup.Enabled = false;
					groupBoxPlanSum.Enabled = true;
					groupBoxPlanSum.Visible = true;
					checkPlanSumPros.Enabled = true;
					checkPlanSumInput.Enabled = true;
					checkPlanSumPros.Enabled = true;
					checkPlanSumInput.Enabled = true;
				}

				else
				{
					groupBoxPlanSum.Enabled = false;
					groupBoxPlanSum.Visible = false;
					checkPlanSumInput.Enabled = false;
					checkPlanSumPros.Enabled = false;
					checkPlanSumPros.Checked = false;
					checkPlanSumInput.Checked = false;
				}
			}

			if (!checkBoxInput.Checked)
			{
				textBoxInput.Text = "";
			}
		}

        public void button1_Click(object sender, EventArgs e)
        {
            string workingLocInternal = @"\\dc3-pr-files\MedPhysics Backup\Data Extractions\Automation\Info Transfer Internal.txt";
			string workingLocExternal = @"\\dc3-pr-files\MedPhysics Backup\Data Extractions\Automation\Info Transfer External.txt";

			List<string> propRun = new List<string>();
			List<string> strucRun = new List<string>();
			List<string> strucPS = new List<string>();
			int propCount = 0;
			int strucCount = 0;
			int psCount = 0;

			File.WriteAllText(workingLocExternal, "");
			File.WriteAllText(workingLocInternal, "");

			if (checkBoxDose.Checked)
			{
				propRun.Add($"{checkBoxDose.Text}");
				propCount++;
			}

			if (checkBoxDVH.Checked)
			{
				propRun.Add($"{checkBoxDVH.Text}");
				propCount++;
			}

			if (checkBoxBrai.Checked)
			{
				strucRun.Add($"{checkBoxBrai.Text}");
				strucCount++;
				if (checkPlanSumBrai.Checked)
				{
					strucPS.Add($"braiPlanSum");
					psCount++;
				}
			}

			if (checkBoxBrea.Checked)
			{
				strucRun.Add($"{checkBoxBrea.Text}");
				strucCount++;
				if (checkPlanSumBrea.Checked)
				{
					strucPS.Add($"breaPlanSum");
					psCount++;
				}
			}

			if (checkBoxLung.Checked)
			{
				strucRun.Add($"{checkBoxLung.Text}");
				strucCount++;
				if (checkPlanSumLung.Checked)
				{
					strucPS.Add($"lungPlanSum");
					psCount++;
				}
			}

			if (checkBoxPros.Checked)
			{
				strucRun.Add($"{checkBoxPros.Text}");
				strucCount++;
				if (checkPlanSumPros.Checked)
				{
					strucPS.Add($"prosPlanSum");
					psCount++;
				}
			}

			if (checkBoxInput.Checked)
			{
				strucRun.Add($"{textBoxInput.Text}");
				strucCount++;
				if (checkPlanSumInput.Checked)
				{
					strucPS.Add($"inputPlanSum");
					psCount++;
				}
			}

			if (comboBox1.Text == "Patient Template Creation")
			{
				if (propRun.Count() > 0)
				{
					if (strucRun.Count() > 0)
					{
						foreach (string property in propRun)
						{
							foreach (string structure in strucRun)
							{
								File.WriteAllText(workingLocExternal, structure);

								if (structure == "Prostate" && strucPS.Contains("prosPlanSum"))
								{
									File.AppendAllText(workingLocExternal, Environment.NewLine + "plansumYes");
								}
								else if (structure != "Input" && !strucPS.Contains("inputPlanSum"))
								{
									File.AppendAllText(workingLocExternal, Environment.NewLine + "plansumNo");
								}

								if (structure == "Input" && strucPS.Contains("inputPlanSum"))
								{
									File.AppendAllText(workingLocExternal, Environment.NewLine + "planSumYes");
								}

								File.AppendAllText(workingLocExternal, Environment.NewLine + $"property{property}");

								var processA = Process.Start(@"\\dc3-pr-files\MedPhysics Backup\Data Extractions\Automation\Executables\1. Patient Template Creation.lnk");
								processA.WaitForExit();
								editingGroup.Enabled = true;
								archiveButton.Enabled = true;
							}
						}
					}
					else
					{
						MessageBox.Show("No Structure Selected.\nPlease make a selection", "Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
					}
				}
				else
				{
					MessageBox.Show("Either DVH or Dose not selected.\nPlease make a selection.", "Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
			}

			else if (comboBox1.Text == "Extraction")
			{
				if (propRun.Count() > 0)
				{
					if (strucRun.Count() > 0)
					{
						foreach (string property in propRun)
						{
							foreach (string structure in strucRun)
							{
								File.WriteAllText(workingLocExternal, structure);

								if (structure == "Prostate" && strucPS.Contains("prosPlanSum"))
								{
									File.AppendAllText(workingLocExternal, Environment.NewLine + "plansumYes");
								}
								else if (structure != "Input" && !strucPS.Contains("inputPlanSum"))
								{
									File.AppendAllText(workingLocExternal, Environment.NewLine + "plansumNo");
								}

								if (structure == "Input" && strucPS.Contains("inputPlanSum"))
								{
									File.AppendAllText(workingLocExternal, Environment.NewLine + "planSumYes");
								}

								File.AppendAllText(workingLocExternal, Environment.NewLine + $"property{property}");

								if (property == "Dose")
								{
									var processA = Process.Start(@"\\dc3-pr-files\MedPhysics Backup\Data Extractions\Automation\Executables\2. Extract Dose.lnk");
									processA.WaitForExit();
								}
								else if (property == "DVH")
								{
									var processA = Process.Start(@"\\dc3-pr-files\MedPhysics Backup\Data Extractions\Automation\Executables\3. Extract DVH.lnk");
									processA.WaitForExit();
								}
							}
						}
					}
					else
					{
						MessageBox.Show("No Structure Selected.\nPlease make a selection", "Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
					}
				}
				else
				{
					MessageBox.Show("Either DVH or Dose not selected.\nPlease make a selection.", "Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
			}

			else if (comboBox1.Text == "Extraction (Single Patient)")
			{
				if (strucRun.Count() > 0)
				{
					if (strucRun.Count() == 1)
					{
						if (textBoxID.TextLength == 8)
						{
							foreach (string property in propRun)
							{
								foreach (string structure in strucRun)
								{
									File.WriteAllText(workingLocExternal, structure);
									File.AppendAllText(workingLocExternal, Environment.NewLine + $"property{property}");

									string inputID = textBoxID.Text;
									File.AppendAllText(workingLocExternal, Environment.NewLine + $"{inputID}");

									if (checkBoxDose.Checked == true)
									{
										var processA = Process.Start(@"\\dc3-pr-files\MedPhysics Backup\Data Extractions\Automation\Executables\2. Extract Dose.lnk");
										processA.WaitForExit();
									}

									if (checkBoxDVH.Checked == true)
									{
										var processA = Process.Start(@"\\dc3-pr-files\MedPhysics Backup\Data Extractions\Automation\Executables\3. Extract DVH.lnk");
										processA.WaitForExit();
									}

									editingGroup.Enabled = true;
									failedButton.Enabled = true;
									archiveButton.Enabled = true;
									spreadsheetButton.Enabled = true;

									File.WriteAllText(workingLocExternal, "");
								}
							}
						}
						else
						{
							MessageBox.Show("Patient ID must be 8 characters in length.", "Invalid Patient ID", MessageBoxButtons.OK, MessageBoxIcon.Error);
						}
					}
					else
					{
						MessageBox.Show("Multiselection of structure not available for Single Patient Extraction", "Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
					}
				}
				else
				{
					MessageBox.Show("No Structure Selected.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				}
			}

			else if (comboBox1.Text == "Run Template and Extraction")
			{
				if (propRun.Count() > 0)
				{
					if (strucRun.Count() > 0)
					{
						foreach (string property in propRun)
						{
							foreach (string structure in strucRun)
							{
								File.WriteAllText(workingLocExternal, structure);

								if (structure == "Prostate" && strucPS.Contains("prosPlanSum"))
								{
									File.AppendAllText(workingLocExternal, Environment.NewLine + "plansumYes");
								}
								else if (structure != "Input" && !strucPS.Contains("inputPlanSum"))
								{
									File.AppendAllText(workingLocExternal, Environment.NewLine + "plansumNo");
								}

								if (structure == "Input" && strucPS.Contains("inputPlanSum"))
								{
									File.AppendAllText(workingLocExternal, Environment.NewLine + "planSumYes");
								}

								File.AppendAllText(workingLocExternal, Environment.NewLine + $"property{property}");

								if (property == "Dose")
								{
									var processA = Process.Start(@"\\dc3-pr-files\MedPhysics Backup\Data Extractions\Automation\Executables\1. Patient Template Creation.lnk");
									processA.WaitForExit();
								}

								else if (property == "DVH")
								{
									var processA = Process.Start(@"\\dc3-pr-files\MedPhysics Backup\Data Extractions\Automation\Executables\1. Patient Template Creation.lnk");
									processA.WaitForExit();
								}

								editingGroup.Enabled = true;
								failedButton.Enabled = true;
								archiveButton.Enabled = true;
								spreadsheetButton.Enabled = true;

								//File.WriteAllText(workingLocExternal, "");
							}
						}

						foreach (string property in propRun)
						{
							foreach (string structure in strucRun)
							{
								File.WriteAllText(workingLocExternal, structure);

								if (structure == "Prostate" && strucPS.Contains("prosPlanSum"))
								{
									File.AppendAllText(workingLocExternal, Environment.NewLine + "plansumYes");
								}
								else if (structure != "Input" && !strucPS.Contains("inputPlanSum"))
								{
									File.AppendAllText(workingLocExternal, Environment.NewLine + "plansumNo");
								}

								if (structure == "Input" && strucPS.Contains("inputPlanSum"))
								{
									File.AppendAllText(workingLocExternal, Environment.NewLine + "planSumYes");
								}

								File.AppendAllText(workingLocExternal, Environment.NewLine + $"property{property}");

								if (property == "Dose")
								{
									var processB = Process.Start(@"\\dc3-pr-files\MedPhysics Backup\Data Extractions\Automation\Executables\2. Extract Dose.lnk");
									processB.WaitForExit();
								}

								else if (property == "DVH")
								{
									var processB = Process.Start(@"\\dc3-pr-files\MedPhysics Backup\Data Extractions\Automation\Executables\3. Extract DVH.lnk");
									processB.WaitForExit();
								}

								editingGroup.Enabled = true;
								failedButton.Enabled = true;
								archiveButton.Enabled = true;
								spreadsheetButton.Enabled = true;

								//File.WriteAllText(workingLocExternal, "");
							}
						}
					}
					else
					{
						MessageBox.Show("No Structure Selected.\nPlease make a selection", "Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
					}
				}
				else
				{
					MessageBox.Show("Either DVH or Dose not selected.\nPlease make a selection.", "Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
			}

			else if (comboBox1.Text == "Clear Templates Archive")
			{
				if (propRun.Count() > 0)
				{
					foreach (string property in propRun)
					{
						File.WriteAllText(workingLocExternal, $"property{property}");

						var processA = Process.Start(@"\\dc3-pr-files\MedPhysics Backup\Data Extractions\Automation\Executables\Clear Templates Archive.lnk");

					}

				}
			}

			else
			{
				MessageBox.Show("Please select an executable to run.", "Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex == 0)
            {
                DialogResult result = MessageBox.Show("Are you sure you want to run Clean Slate Protocol?", "Clean Slate Protocol",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (result == DialogResult.Yes)
                {
                    var cleanSlateProt = Process.Start(@"\\dc3-pr-files\MedPhysics Backup\Data Extractions\Automation\Executables\Clean Slate Protocol (TESTING PURPOSES ONLY).lnk");
                    cleanSlateProt.WaitForExit();
                }
                else if (result == DialogResult.No)
                {
                    MessageBox.Show("Protocol Aborted.", "Clean Slate Protocol", MessageBoxButtons.OK);
                }
            }
        }

		private void button4_Click(object sender, EventArgs e)
		{
			List<string> endLocList = new List<string>();
			List<string> propList = new List<string>();

			if (checkBoxPros.Checked == true)
			{
				endLocList.Add(checkBoxPros.Text);
			}

			if (checkBoxLung.Checked == true)
			{
				endLocList.Add(checkBoxLung.Text);
			}

			if (checkBoxBrea.Checked == true)
			{
				endLocList.Add(checkBoxBrea.Text);
			}

			if (checkBoxBrai.Checked == true)
			{
				endLocList.Add(checkBoxBrai.Text);
			}

			if (checkBoxInput.Checked == true)
			{
				endLocList.Add(textBoxInput.Text);
			}

			if (checkBoxDose.Checked)
			{
				propList.Add("Dose");
			}

			if (checkBoxDVH.Checked)
			{
				propList.Add("DVH");
			}

			if (propList.Count() > 0)
			{
				if (endLocList.Count() > 0)
				{
					foreach (string property in propList)
					{
						foreach (string endLoc in endLocList)
						{
							string inputInfotxt = $@"\\dc3-pr-files\MedPhysics Backup\Data Extractions\Automation\{property}\Input Information Text Files\{endLoc} {property} Input Information.txt";

							if (File.Exists(inputInfotxt))
							{
								System.Diagnostics.Process.Start(inputInfotxt);
							}
							else
							{
								if (endLoc != "")
								{
									DialogResult result = MessageBox.Show($"No input information file exists for {endLoc} {property}. \nWould you like to open a new file?", "File Does Not Exist",
										MessageBoxButtons.YesNo, MessageBoxIcon.Error);
									if (result == DialogResult.Yes)
									{
										inputInfotxt = $@"\\dc3-pr-files\MedPhysics Backup\Data Extractions\Automation\{property}\Input Information Text Files\{endLoc} {property} Input Information.txt";
										string genInputInfotxt = $@"\\dc3-pr-files\MedPhysics Backup\Data Extractions\Automation\{property}\Input Information Text Files\General {property} Input Information.txt";
										File.Copy(genInputInfotxt, inputInfotxt);
										System.Diagnostics.Process.Start(inputInfotxt);
									}
								}
							}
						}
					}
				}
				else
				{
					MessageBox.Show("No Structure Selected.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				}
			}
			else
			{
				MessageBox.Show("Either dose or DVH not selected. \nPlease make a selection.", "Extraction Information", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void button5_Click(object sender, EventArgs e)
        {
			List<string> condList = new List<string>();
			List<string> propList = new List<string>();

            if (checkBoxPros.Checked)
            {
                condList.Add(checkBoxPros.Text);
            }

            if (checkBoxLung.Checked)
            {
                condList.Add(checkBoxLung.Text);
            }

            if (checkBoxBrea.Checked)
            {
                condList.Add(checkBoxBrea.Text);
            }

            if (checkBoxBrai.Checked)
            {
                condList.Add(checkBoxBrai.Text);
            }

            if (checkBoxInput.Checked)
            {
                condList.Add(textBoxInput.Text);
            }

			if (checkBoxDose.Checked)
			{
				propList.Add("Dose");
			}
			if (checkBoxDVH.Checked)
			{
				propList.Add("DVH");
			}

			if (propList.Count() > 0)
			{
				if (condList.Count() > 0)
				{
					foreach (string property in propList)
					{
						foreach (string cond in condList)
						{
							string inputInfotxt = $@"\\dc3-pr-files\MedPhysics Backup\Data Extractions\Automation\{property}\Input Information Text Files\{cond} {property} Input Information.txt";
							string condFile = "";

							if (File.Exists(inputInfotxt))
							{
								condFile = File.ReadLines(inputInfotxt).Skip(4).Take(1).First();

								if (File.Exists(condFile))
								{
									System.Diagnostics.Process.Start(condFile);
								}
								else
								{
									if (cond != "")
									{
										DialogResult result = MessageBox.Show($"Conditional file for {cond} {property} does not exist. \nWould you like to open the blank conditional template?", "File Does Not Exist",
											MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
										if (result == DialogResult.Yes)
										{
											string genCondFile = $@"\\dc3-pr-files\MedPhysics Backup\Data Extractions\Automation\{property}\Template Creation Input\Conditional Template.xlsx";
											System.Diagnostics.Process.Start(genCondFile);
										}
									}
								}
							}
							else
							{
								if (cond != "")
								{
									DialogResult result = MessageBox.Show($"No input information file exists for {cond} {property}. \nWould you like to open the general file?", "File Does Not Exist",
										MessageBoxButtons.YesNo, MessageBoxIcon.Error);
									if (result == DialogResult.Yes)
									{
										inputInfotxt = $@"\\dc3-pr-files\MedPhysics Backup\Data Extractions\Automation\{property}\Input Information Text Files\General {property} Input Information.txt";
										System.Diagnostics.Process.Start(inputInfotxt);
									}
								}
							}
						}
					}
				}
				else
				{
					MessageBox.Show("No Structure Selected.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				}
			}
			else
			{
				MessageBox.Show("Either dose or DVH not selected.\nPlease make a selection.", "Extraction Information", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void progressBar1_Click(object sender, EventArgs e)
        {
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {
			string inputInfotxt = "";

			List<string> propList = new List<string>();

			if (checkBoxDose.Checked)
			{
				propList.Add("Dose");
			}

			if (checkBoxDVH.Checked)
			{
				propList.Add("DVH");
			}

			if (propList.Count() > 0)
			{
				foreach (string property in propList)
				{
					inputInfotxt = $@"\\dc3-pr-files\MedPhysics Backup\Data Extractions\Automation\{property}\Input Information Text Files\General {property} Input Information.txt";
					if (File.Exists(inputInfotxt))
					{
						string archiveLoc = File.ReadLines(inputInfotxt).Skip(20).Take(1).First();
						System.Diagnostics.Process.Start(archiveLoc);
					}
				}
			}
			else
			{
				MessageBox.Show("Either dose or DVH not selected.\nPlease make a selection.", "Extraction Information", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
        }

        private void button7_Click(object sender, EventArgs e)
        {
			List<string> condList = new List<string>();
			List<string> propList = new List<string>();

            if (checkBoxPros.Checked)
            {
                condList.Add(checkBoxPros.Text);
            }

            if (checkBoxLung.Checked)
            {
                condList.Add(checkBoxLung.Text);
            }

            if (checkBoxBrea.Checked)
            {
                condList.Add(checkBoxBrea.Text);
            }

            if (checkBoxBrai.Checked)
            {
                condList.Add(checkBoxBrai.Text);
            }

            if (checkBoxInput.Checked)
            {
                condList.Add(textBoxInput.Text);
            }


			if (checkBoxDose.Checked)
			{
				propList.Add("Dose");
			}

			if (checkBoxDVH.Checked)
			{
				propList.Add("DVH");
			}

			if (propList.Count() > 0)
			{
				if (condList.Count() > 0)
				{
					foreach (string property in propList)
					{
						foreach (string cond in condList)
						{
							string inputInfotxt = $@"\\dc3-pr-files\MedPhysics Backup\Data Extractions\Automation\{property}\Input Information Text Files\{cond} {property} Input Information.txt";
							string patientInfo = "";
							if (File.Exists(inputInfotxt))
							{
								patientInfo = File.ReadLines(inputInfotxt).Skip(8).Take(1).First();

								if (File.Exists(patientInfo))
								{
									System.Diagnostics.Process.Start(patientInfo);
								}
								else
								{
									if (cond != "")
									{
										DialogResult result = MessageBox.Show($"Template creation input structure, {cond}, not found for {property}. \nWould you like to open Template Creation Input Folder?", "Template Creation Input",
											MessageBoxButtons.YesNo, MessageBoxIcon.Error);
										if (result == DialogResult.Yes)
										{
											string creationInputLoc = $@"\\dc3-pr-files\MedPhysics Backup\Data Extractions\Automation\{property}\Template Creation Input\";
											System.Diagnostics.Process.Start(creationInputLoc);
										}
									}
								}
							}
							else
							{
								if (cond != "")
								{
									DialogResult result = MessageBox.Show($"No input information file exists for {cond} {property}. \nWould you like to open the general file?", "File Does Not Exist",
										MessageBoxButtons.YesNo, MessageBoxIcon.Error);
									if (result == DialogResult.Yes)
									{
										inputInfotxt = $@"\\dc3-pr-files\MedPhysics Backup\Data Extractions\Automation\{property}\Input Information Text Files\General {property} Input Information.txt";
										System.Diagnostics.Process.Start(inputInfotxt);
									}
								}
							}
						}
					}
				}
				else
				{
					MessageBox.Show("No Structure Selected.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				}
			}
			else
			{
				MessageBox.Show("Either dose or DVH not selected.\nPlease make a selection.", "Extraction Information", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void button8_Click(object sender, EventArgs e)
        {

			List<string> propList = new List<string>();
			List<string> strucList = new List<string>();

            if (checkBoxPros.Checked)
            {
                strucList.Add(checkBoxPros.Text);
            }

            if (checkBoxLung.Checked)
            {
                strucList.Add(checkBoxLung.Text);
            }

            if (checkBoxBrea.Checked)
            {
                strucList.Add(checkBoxBrea.Text);
            }

            if (checkBoxBrai.Checked)
            {
                strucList.Add(checkBoxBrai.Text);
            }

            if (checkBoxInput.Checked)
            {
                strucList.Add(textBoxInput.Text);
            }

			if (checkBoxDose.Checked)
			{
				propList.Add("Dose");
			}

			if (checkBoxDVH.Checked)
			{
				propList.Add("DVH");
			}

			if (propList.Count() > 0)
			{
				if (strucList.Count() > 0)
				{
					foreach (string property in propList)
					{
						foreach (string structure in strucList)
						{
							string inputInfotxt = $@"\\dc3-pr-files\MedPhysics Backup\Data Extractions\Automation\{property}\Input Information Text Files\{structure} {property} Input Information.txt";
							string patientList = "";
							if (File.Exists(inputInfotxt))
							{
								patientList = File.ReadLines(inputInfotxt).Skip(32).Take(1).First();

								if (File.Exists(patientList))
								{
									System.Diagnostics.Process.Start(patientList);
								}
								else
								{
									if (structure != "")
									{
										DialogResult result = MessageBox.Show($"No Available Patient ID List for {structure} in {property}. \nWould you like to open Patient ID Lists Folder?", "Patient ID Lists",
											MessageBoxButtons.YesNo, MessageBoxIcon.Error);
										if (result == DialogResult.Yes)
										{
											string patientListLoc = $@"\\dc3-pr-files\MedPhysics Backup\Data Extractions\Automation\{property}\Patient ID Lists\";
											System.Diagnostics.Process.Start(patientListLoc);
										}
									}
								}
							}
							else
							{
								if (structure != "")
								{
									DialogResult result = MessageBox.Show($"No input information file exists for {structure} {property}. \nWould you like to open the general file?", "File Does Not Exist",
										MessageBoxButtons.YesNo, MessageBoxIcon.Error);
									if (result == DialogResult.Yes)
									{
										inputInfotxt = $@"\\dc3-pr-files\MedPhysics Backup\Data Extractions\Automation\{property}\Input Information Text Files\General {property} Input Information.txt";
										System.Diagnostics.Process.Start(inputInfotxt);
									}
								}
							}
						}
					}
				}
				else
				{
					MessageBox.Show("No Structure Selected.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				}
			}
			else
			{
				MessageBox.Show("Either dose or DVH not selected.\nPlease make a selection.", "Extraction Information", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

        private void button9_Click(object sender, EventArgs e)
        {
			List<string> propList = new List<string>();
			List<string> strucList = new List<string>();

            if (checkBoxPros.Checked == true)
            {
                strucList.Add(checkBoxPros.Text);
            }

            if (checkBoxLung.Checked == true)
            {
                strucList.Add(checkBoxLung.Text);
            }

            if (checkBoxBrea.Checked == true)
            {
                strucList.Add(checkBoxBrea.Text);
            }

            if (checkBoxBrai.Checked == true)
            {
                strucList.Add(checkBoxBrai.Text);
            }

            if (checkBoxInput.Checked == true)
            {
                strucList.Add(textBoxInput.Text);
            }


			if (checkBoxDose.Checked == true)
			{
				propList.Add("Dose");
			}

			if (checkBoxDVH.Checked == true)
			{
				propList.Add("DVH");
			}

			if (propList.Count() > 0)
			{
				if (strucList.Count() > 0)
				{
					foreach (string property in propList)
					{
						foreach (string structure in strucList)
						{
							string inputInfotxt = $@"\\dc3-pr-files\MedPhysics Backup\Data Extractions\Automation\{property}\Input Information Text Files\{structure} {property} Input Information.txt";
							string templateLoc = "";
							if (File.Exists(inputInfotxt))
							{
								templateLoc = File.ReadLines(inputInfotxt).Skip(16).Take(1).First();

								if (Directory.Exists(templateLoc))
								{
									System.Diagnostics.Process.Start(templateLoc);
								}
								else
								{
									DialogResult result = MessageBox.Show($"Templates for {structure} do not exist. \nWould you like to open the Template folder location?", "Template Location",
										MessageBoxButtons.YesNo, MessageBoxIcon.Error);
									if (result == DialogResult.Yes)
									{
										string templateFolderLoc = $@"\\dc3-pr-files\MedPhysics Backup\Data Extractions\Automation\{property}\Template Creation Output\";
										System.Diagnostics.Process.Start(templateFolderLoc);
									}
								}
							}
							else
							{
								DialogResult result = MessageBox.Show($"No input information file exists for {structure} {property}. \nWould you like to open the general file?", "File Does Not Exist",
									MessageBoxButtons.YesNo, MessageBoxIcon.Error);
								if (result == DialogResult.Yes)
								{
									inputInfotxt = $@"\\dc3-pr-files\MedPhysics Backup\Data Extractions\Automation\{property}\Input Information Text Files\General {property} Input Information.txt";
									System.Diagnostics.Process.Start(inputInfotxt);
								}
							}
						}
					}
				}
				else
				{
					MessageBox.Show("No Structure Selected.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				}
			}
			else
			{
				MessageBox.Show("Either dose or DVH not selected.\nPlease make a selection.", "Extraction Information", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void button10_Click(object sender, EventArgs e)
        {
			List<string> propList = new List<string>();
			List<string> strucList = new List<string>();

            if (checkBoxPros.Checked)
            {
                strucList.Add(checkBoxPros.Text);
            }

            if (checkBoxLung.Checked)
            {
                strucList.Add(checkBoxLung.Text);
            }

            if (checkBoxBrea.Checked)
            {
                strucList.Add(checkBoxBrea.Text);
            }

            if (checkBoxBrai.Checked)
            {
                strucList.Add(checkBoxBrai.Text);
            }
            
            if (checkBoxInput.Checked)
            {
                strucList.Add(textBoxInput.Text);
            }


			if (checkBoxDose.Checked)
			{
				propList.Add("Dose");
			}

			if (checkBoxDVH.Checked)
			{
				propList.Add("DVH");
			}

			if (propList.Count() > 0)
			{
				if (strucList.Count() > 0)
				{
					foreach (string property in propList)
					{
						foreach (string structure in strucList)
						{
							string inputInfotxt = $@"\\dc3-pr-files\MedPhysics Backup\Data Extractions\Automation\{property}\Input Information Text Files\{structure} {property} Input Information.txt";
							string failedLoc = "";
							if (File.Exists(inputInfotxt))
							{
								failedLoc = File.ReadLines(inputInfotxt).Skip(36).Take(1).First();

								if (File.Exists(failedLoc))
								{
									System.Diagnostics.Process.Start(failedLoc);
								}
								else
								{
									if (structure != "")
									{
										DialogResult result = MessageBox.Show($"Could not find failed extraction text file for {structure} in {property}. \nWould you like to open {property} folder?", "Failed Extraction Folder",
											MessageBoxButtons.YesNo, MessageBoxIcon.Error);
										if (result == DialogResult.Yes)
										{
											string automationLoc = $@"\\dc3-pr-files\MedPhysics Backup\Data Extractions\Automation\{property}\";
											System.Diagnostics.Process.Start(automationLoc);
										}
									}
								}
							}
							else
							{
								if (structure != "")
								{
									DialogResult result = MessageBox.Show($"No input information file exists for {structure} {property}. \nWould you like to open the general file?", "File Does Not Exist",
										MessageBoxButtons.YesNo, MessageBoxIcon.Error);
									if (result == DialogResult.Yes)
									{
										inputInfotxt = $@"\\dc3-pr-files\MedPhysics Backup\Data Extractions\Automation\{property}\Input Information Text Files\General {property} Input Information.txt";
										System.Diagnostics.Process.Start(inputInfotxt);
									}
								}
							}
						}
					}
				}
				else
				{
					MessageBox.Show("No Structure Selected.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				}
			}
			else
			{
				MessageBox.Show("Either dose or DVH not selected.\nPlease make a selection.", "Extraction Information", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button11_Click(object sender, EventArgs e)
        {
			List<string> propList = new List<string>();
			List<string> strucList = new List<string>();

            if (checkBoxPros.Checked)
            {
                strucList.Add(checkBoxPros.Text);
            }

            if (checkBoxLung.Checked)
            {
                strucList.Add(checkBoxLung.Text);
            }

            if (checkBoxBrea.Checked)
            {
                strucList.Add(checkBoxBrea.Text);
            }

            if (checkBoxBrai.Checked)
            {
                strucList.Add(checkBoxBrai.Text);
            }

            if (checkBoxInput.Checked)
            {
                strucList.Add(textBoxInput.Text);
            }

			if (checkBoxDose.Checked)
			{
				propList.Add("Dose");
			}

			if (checkBoxDVH.Checked)
			{
				propList.Add("DVH");
			}


			if (propList.Count() > 0)
			{
				if (strucList.Count() > 0)
				{
					foreach (string property in propList)
					{
						foreach (string structure in strucList)
						{
							string inputInfotxt = $@"\\dc3-pr-files\MedPhysics Backup\Data Extractions\Automation\{property}\Input Information Text Files\{structure} {property} Input Information.txt";
							string spreadsheet = "";
							if (File.Exists(inputInfotxt))
							{
								spreadsheet = File.ReadLines(inputInfotxt).Skip(40).Take(1).First();

								if (File.Exists(spreadsheet))
								{
									System.Diagnostics.Process.Start(spreadsheet);
								}
								else
								{
									if (structure != "")
									{
										DialogResult result = MessageBox.Show($"Could not find patient spreadsheet for {structure} in {property}. \nWould you like to open Extracted {property} folder?", "Patient Spreadsheet Excel File",
											MessageBoxButtons.YesNo, MessageBoxIcon.Error);
										if (result == DialogResult.Yes)
										{
											string spreadsheetLoc = $@"\\dc3-pr-files\MedPhysics Backup\Data Extractions\Extracted {property}";
											System.Diagnostics.Process.Start(spreadsheetLoc);
										}
									}
								}
							}
							else
							{
								if (structure != "")
								{
									DialogResult result = MessageBox.Show($"No input information file exists for {structure} {property}. \nWould you like to open the general file?", "File Does Not Exist",
										MessageBoxButtons.YesNo, MessageBoxIcon.Error);
									if (result == DialogResult.Yes)
									{
										inputInfotxt = $@"\\dc3-pr-files\MedPhysics Backup\Data Extractions\Automation\{property}\Input Information Text Files\General {property} Input Information.txt";
										System.Diagnostics.Process.Start(inputInfotxt);
									}
								}
							}
						}
					}
				}
				else
				{
					MessageBox.Show("No Structure Selected.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				}
			}
			else
			{
				MessageBox.Show("Either dose or DVH not selected.\nPlease make a selection.", "Extraction Information", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

        private void button1_Click_1(object sender, EventArgs e)
        {

            comboBox1.SelectedIndex = 0;

            userInputGroup.Enabled = true;
            editingGroup.Enabled = true;
            archiveButton.Enabled = true;
            failedButton.Enabled = true;
            spreadsheetButton.Enabled = true;

            groupBoxStructure.Enabled = true;
            checkBoxPros.Enabled = true;
            checkBoxBrai.Checked = true;
            checkBoxLung.Enabled = true;
            checkBoxBrea.Enabled = true;
            checkBoxBrai.Enabled = true;
            checkBoxInput.Enabled = true;
            textBoxInput.Enabled = true;
            textBoxInput.Text = "";

            groupBoxPlanSum.Enabled = false;
            checkPlanSumInput.Enabled = false;
            checkPlanSumPros.Enabled = false;
            checkPlanSumInput.Checked = false;
            checkPlanSumPros.Checked = false;

            runExeButton.Enabled = false;

            cleanSlateButton.Enabled = true;

			checkBoxDose.Enabled = true;
			checkBoxDVH.Enabled = true;
			checkBoxDose.Checked = true;
        }

        private void radioButton2_CheckedChanged_1(object sender, EventArgs e)
        {

        }

        private void radioButton1_CheckedChanged_1(object sender, EventArgs e)
        {

        }

        private void groupBox1_Enter_1(object sender, EventArgs e)
        {

        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }

        private void editingGroup_Enter(object sender, EventArgs e)
        {

        }

        private void radioButton3_CheckedChanged_1(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged_2(object sender, EventArgs e)
        {
            
        }

        private void feedbackButton_Click(object sender, EventArgs e)
        {
            DateTime localDate = DateTime.Now;
            string localDateString = localDate.ToString("yyyy-MM-dd   hh mm tt");

            string feedbackLoc = $@"\\dc3-pr-files\MedPhysics Backup\Data Extractions\Automation\Feedback\{localDateString}.txt";
            if (!File.Exists(feedbackLoc))
            {
                File.Create(feedbackLoc).Dispose();
                //System.Diagnostics.Process.Start(feedbackLoc);
                var processFeedback = Process.Start(feedbackLoc);
                processFeedback.WaitForExit();
                if (File.ReadAllLines(feedbackLoc).Length == 0)
                {
                    File.Delete(feedbackLoc);
                }

            }
            else if (File.Exists(feedbackLoc))
            {
                //System.Diagnostics.Process.Start(feedbackLoc);
                var processFeedback = Process.Start(feedbackLoc);
                processFeedback.WaitForExit();
                if (File.ReadAllLines(feedbackLoc).Length == 0)
                {
                    File.Delete(feedbackLoc);
                }
            }
        }

        private void textBoxID_TextChanged(object sender, EventArgs e)
        {
			//string inputInfoPros = "";
			//string inputInfoBrea = "";
			//string inputInfoLung = "";
			//string inputInfoBrai = "";
			string property = "";

			if (comboBox1.Text == "Extract DVH" || comboBox1.Text == "Extract DVH (Single Patient)")
			{
				property = "DVH";
			}
			else
			{
				property = "Dose";
			}

            string inputInfoPros = $@"\\dc3-pr-files\MedPhysics Backup\Data Extractions\Automation\Dose\Input Information Text Files\Prostate {property} Input Information.txt";
            string inputInfoBrea = $@"\\dc3-pr-files\MedPhysics Backup\Data Extractions\Automation\Dose\Input Information Text Files\Breast {property} Input Information.txt";
            string inputInfoLung = $@"\\dc3-pr-files\MedPhysics Backup\Data Extractions\Automation\Dose\Input Information Text Files\Lung {property} Input Information.txt";
            string inputInfoBrai = $@"\\dc3-pr-files\MedPhysics Backup\Data Extractions\Automation\Dose\Input Information Text Files\Brain {property} Input Information.txt";

            string patientListPros = File.ReadLines(inputInfoPros).Skip(32).Take(1).First();
            string patientListBrea = File.ReadLines(inputInfoBrea).Skip(32).Take(1).First();
            string patientListLung = File.ReadLines(inputInfoLung).Skip(32).Take(1).First();
            string patientListBrai = File.ReadLines(inputInfoBrai).Skip(32).Take(1).First();

            if (textBoxID.Text != "")
            {
                labelID.Visible = false;

                if (File.ReadAllLines(patientListPros).Contains(textBoxID.Text))
                {
                    checkBoxPros.Checked = true;
                }
                else if (File.ReadAllLines(patientListBrea).Contains(textBoxID.Text))
                {
                    checkBoxBrea.Checked = true;
                }
                else if (File.ReadAllLines(patientListLung).Contains(textBoxID.Text))
                {
                    checkBoxLung.Checked = true;
                }
                else if (File.ReadAllLines(patientListBrai).Contains(textBoxID.Text))
                {
                    checkBoxBrai.Checked = true;
                }
                else
                {
                    checkBoxPros.Checked = false;
                    checkBoxBrea.Checked = false;
                    checkBoxLung.Checked = false;
                    checkBoxBrai.Checked = false;
                }
                
                if (textBoxID.TextLength == 8)
                {
                    singleTemplateButton.Enabled = true;
                    singleTemplateButton.Visible = true;

                    groupBoxStructure.Enabled = true;
                    userInputGroup.Enabled = true;
                    editingGroup.Enabled = true;
                }
                else
                {
                    singleTemplateButton.Enabled = false;
                    singleTemplateButton.Visible = false;

                    groupBoxStructure.Enabled = false;
                    userInputGroup.Enabled = false;
                    editingGroup.Enabled = false;
                }
            }
        }

        private void labelID_Click(object sender, EventArgs e)
        {

        }

        private void singleTemplateButton_Click(object sender, EventArgs e)
        {
            string singleID = textBoxID.Text;

			List<string> strucList = new List<string>();
			List<string> propList = new List<string>();

            if (singleID.Length == 8)
            {
                if (checkBoxPros.Checked)
                {
                    strucList.Add(checkBoxPros.Text);
                }

                if (checkBoxBrea.Checked)
                {
                    strucList.Add(checkBoxBrea.Text);
                }

                if (checkBoxLung.Checked)
                {
                    strucList.Add(checkBoxLung.Text);
                }

                if (checkBoxBrai.Checked)
                {
                    strucList.Add(checkBoxBrai.Text);
                }

                if (checkBoxInput.Checked)
                {
                    strucList.Add(textBoxInput.Text);
                }

				if (checkBoxDose.Checked)
				{
					propList.Add("Dose");
				}

				if (checkBoxDVH.Checked)
				{
					propList.Add("DVH");
				}

				if (propList.Count() > 0)
				{
					if (strucList.Count() > 0)
					{
						foreach (string property in propList)
						{
							foreach (string structure in strucList)
							{
								string inputInfoTxt = $@"\\dc3-pr-files\MedPhysics Backup\Data Extractions\Automation\{property}\Input Information Text Files\{structure} {property} Input Information.txt";

								if (File.Exists(inputInfoTxt))
								{
									string templateLoc = File.ReadLines(inputInfoTxt).Skip(16).Take(1).First();
									string singleTemplateLoc = $@"{templateLoc}\{singleID}.txt";
									if (File.Exists(singleTemplateLoc))
									{
										System.Diagnostics.Process.Start(singleTemplateLoc);
									}
									else
									{
										if (structure != "")
										{
											DialogResult result = MessageBox.Show($"Patient, {singleID}, does not have a template. \nWould you like to open the {structure} templates?", "Template Does Not Exist",
												MessageBoxButtons.YesNo, MessageBoxIcon.Question);
											if (result == DialogResult.Yes)
											{
												System.Diagnostics.Process.Start(templateLoc);
											}
										}
									}
								}
								else
								{
									if (structure != "")
									{
										DialogResult result = MessageBox.Show($"Input info for {structure} in {property} does not exist. \nWould you like to open the Input Info folder location?", "Input Info Location",
											MessageBoxButtons.YesNo, MessageBoxIcon.Error);
										if (result == DialogResult.Yes)
										{
											string templateFolderLoc = $@"\\dc3-pr-files\MedPhysics Backup\Data Extractions\Automation\{property}\Input Information Text Files\";
											System.Diagnostics.Process.Start(templateFolderLoc);
										}
									}
								}
							}
						}
					}
					else
					{
						MessageBox.Show("No Structure Selected.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
					}
				}
				else
				{
					MessageBox.Show("Either dose or DVH not selected.\nPlease make a selection.", "Extraction Information", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
			}
			else
			{
				MessageBox.Show("Patient ID must be 8 characters in length.\nPlease either change ID in plan or change input.", "Patient ID", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
        }

		private void radioButtonDose_CheckedChanged(object sender, EventArgs e)
		{

		}

		private void radioButtonDVH_CheckedChanged(object sender, EventArgs e)
		{

		}

		private void buttonQuery_Click(object sender, EventArgs e)
		{
			string exePath = @"\\dc3-pr-files\MedPhysics Backup\Coop Students\2019\Term 3 - Autumn\Quinton Tennant\Release Copy\SQL Script Calling GUI\SQL Script Calling GUI\bin\Debug\SQL Script Calling GUI.exe";

			Process process = new Process();
			// Pass your exe file path here.
			string path = exePath;
			string fileName = Path.GetFileName(path);
			// Get the precess that already running as per the exe file name.
			Process[] processName = Process.GetProcessesByName(fileName.Substring(0, fileName.LastIndexOf('.')));
			if (processName.Length > 0)
			{
				MessageBox.Show("Instance of Database Query already open", "exe Already Open", MessageBoxButtons.OK, MessageBoxIcon.Warning);
			}
			else
			{
				string queryExe = @"\\dc3-pr-files\MedPhysics Backup\Data Extractions\Automation\Executables\Database Query GUI.lnk";
				System.Diagnostics.Process.Start(queryExe);
			}
		}

		private void checkBoxDose_CheckedChanged(object sender, EventArgs e)
		{
			if (comboBox1.Text == "Patient Template Creation")
			{
				editingGroup.Enabled = false;
				userInputGroup.Enabled = true;
				if (checkBoxDose.Checked || checkBoxDVH.Checked)
				{
					groupBoxPlanSum.Enabled = true;
					groupBoxPlanSum.Visible = true;
					checkPlanSumPros.Enabled = true;
					checkPlanSumInput.Enabled = true;
				}
				else
				{
					groupBoxPlanSum.Enabled = false;
					groupBoxPlanSum.Visible = true;
					checkPlanSumPros.Enabled = false;
					checkPlanSumInput.Enabled = false;
				}
			}
			else if (comboBox1.Text == "Extraction" || comboBox1.Text == "Extraction (Single Patient)")
			{
				editingGroup.Enabled = true;
				userInputGroup.Enabled = true;

				groupBoxPlanSum.Enabled = false;
				groupBoxPlanSum.Visible = false;
				checkPlanSumPros.Enabled = false;
				checkPlanSumInput.Enabled = false;
			}
			else if (comboBox1.Text == "Run Template and Extraction")
			{
				editingGroup.Enabled = true;
				userInputGroup.Enabled = true;

				groupBoxPlanSum.Enabled = true;
				groupBoxPlanSum.Visible = true;
				checkPlanSumPros.Enabled = true;
				checkPlanSumInput.Enabled = true;
			}
		}

		private void checkBoxDVH_CheckedChanged(object sender, EventArgs e)
		{
			if (comboBox1.Text == "Patient Template Creation")
			{
				editingGroup.Enabled = false;
				userInputGroup.Enabled = true;
				if (checkBoxDose.Checked || checkBoxDVH.Checked)
				{
					groupBoxPlanSum.Enabled = true;
					groupBoxPlanSum.Visible = true;
					checkPlanSumPros.Enabled = true;
					checkPlanSumInput.Enabled = true;
				}
				else
				{
					groupBoxPlanSum.Enabled = false;
					groupBoxPlanSum.Visible = true;
					checkPlanSumPros.Enabled = false;
					checkPlanSumInput.Enabled = false;
				}

			}
			else if (comboBox1.Text == "Extraction" || comboBox1.Text == "Extraction (Single Patient)")
			{
				editingGroup.Enabled = true;
				userInputGroup.Enabled = true;

				groupBoxPlanSum.Enabled = false;
				groupBoxPlanSum.Visible = false;
				checkPlanSumPros.Enabled = false;
				checkPlanSumInput.Enabled = false;
			}
			else if (comboBox1.Text == "Run Template and Extraction")
			{
				editingGroup.Enabled = true;
				userInputGroup.Enabled = true;

				groupBoxPlanSum.Enabled = true;
				groupBoxPlanSum.Visible = true;
				checkPlanSumPros.Enabled = true;
				checkPlanSumInput.Enabled = true;
			}
		}

		private void checkedListBox_SelectedIndexChanged(object sender, EventArgs e)
		{

		}

		private void toolTip1_Popup(object sender, PopupEventArgs e)
		{

		}

		private void toolTipBrai_Popup(object sender, PopupEventArgs e)
		{
		}

		private void checkPlanSumBrai_CheckedChanged(object sender, EventArgs e)
		{

		}

		private void linkLabelSelAll_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
		{
			checkBoxBrai.Checked = true;
			checkBoxBrea.Checked = true;
			checkBoxLung.Checked = true;
			checkBoxPros.Checked = true;

			if (comboBox1.Text == "Patient Template Creation" || comboBox1.Text == "Run Template and Extraction")
			{
				checkPlanSumPros.Checked = true;
			}

			linkLabelSelAll.Enabled = false;
			linkLabelSelAll.Visible = false;

			linkLabelDesAll.Enabled = true;
			linkLabelDesAll.Visible = true;
		}

		private void linkLabelDesAll_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
		{
			checkBoxBrai.Checked = false;
			checkBoxBrea.Checked = false;
			checkBoxLung.Checked = false;
			checkBoxPros.Checked = false;
			checkBoxInput.Checked = false;

			checkPlanSumPros.Checked = false;
			checkPlanSumBrai.Checked = false;
			checkPlanSumBrea.Checked = false;
			checkPlanSumInput.Checked = false;
			checkPlanSumLung.Checked = false;

			linkLabelDesAll.Enabled = false;
			linkLabelDesAll.Visible = false;

			linkLabelSelAll.Enabled = true;
			linkLabelSelAll.Visible = true;
		}

		private void checkPlanSumInput_CheckedChanged(object sender, EventArgs e)
		{

		}

		private void checkPlanSumPros_CheckedChanged(object sender, EventArgs e)
		{

		}
	}
}
