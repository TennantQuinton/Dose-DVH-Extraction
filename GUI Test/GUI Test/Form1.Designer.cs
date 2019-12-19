namespace GUI_Test
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.readmeButton = new System.Windows.Forms.Button();
            this.runExeButton = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.cleanSlateButton = new System.Windows.Forms.Button();
            this.inputInfoButton = new System.Windows.Forms.Button();
            this.userInputGroup = new System.Windows.Forms.GroupBox();
            this.patientInfoButton = new System.Windows.Forms.Button();
            this.conditionalButton = new System.Windows.Forms.Button();
            this.archiveButton = new System.Windows.Forms.Button();
            this.editingGroup = new System.Windows.Forms.GroupBox();
            this.singleTemplateButton = new System.Windows.Forms.Button();
            this.templatesButton = new System.Windows.Forms.Button();
            this.idListButton = new System.Windows.Forms.Button();
            this.failedButton = new System.Windows.Forms.Button();
            this.spreadsheetButton = new System.Windows.Forms.Button();
            this.overrideButton = new System.Windows.Forms.Button();
            this.groupBoxStructure = new System.Windows.Forms.GroupBox();
            this.linkLabelDesAll = new System.Windows.Forms.LinkLabel();
            this.linkLabelSelAll = new System.Windows.Forms.LinkLabel();
            this.groupBoxPlanSum = new System.Windows.Forms.GroupBox();
            this.label3 = new System.Windows.Forms.Label();
            this.checkPlanSumBrai = new System.Windows.Forms.CheckBox();
            this.checkPlanSumLung = new System.Windows.Forms.CheckBox();
            this.checkPlanSumBrea = new System.Windows.Forms.CheckBox();
            this.checkPlanSumInput = new System.Windows.Forms.CheckBox();
            this.checkPlanSumPros = new System.Windows.Forms.CheckBox();
            this.checkBoxInput = new System.Windows.Forms.CheckBox();
            this.checkBoxPros = new System.Windows.Forms.CheckBox();
            this.checkBoxLung = new System.Windows.Forms.CheckBox();
            this.checkBoxBrea = new System.Windows.Forms.CheckBox();
            this.checkBoxBrai = new System.Windows.Forms.CheckBox();
            this.textBoxInput = new System.Windows.Forms.TextBox();
            this.feedbackButton = new System.Windows.Forms.Button();
            this.textBoxID = new System.Windows.Forms.TextBox();
            this.labelID = new System.Windows.Forms.Label();
            this.buttonQuery = new System.Windows.Forms.Button();
            this.checkBoxDose = new System.Windows.Forms.CheckBox();
            this.checkBoxDVH = new System.Windows.Forms.CheckBox();
            this.toolTipBrai = new System.Windows.Forms.ToolTip(this.components);
            this.toolTipBrea = new System.Windows.Forms.ToolTip(this.components);
            this.toolTipLung = new System.Windows.Forms.ToolTip(this.components);
            this.userInputGroup.SuspendLayout();
            this.editingGroup.SuspendLayout();
            this.groupBoxStructure.SuspendLayout();
            this.groupBoxPlanSum.SuspendLayout();
            this.SuspendLayout();
            // 
            // comboBox1
            // 
            this.comboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.IntegralHeight = false;
            this.comboBox1.ItemHeight = 13;
            this.comboBox1.Items.AddRange(new object[] {
            " ",
            "Patient Template Creation",
            "Extraction",
            "Extraction (Single Patient)",
            "Run Template and Extraction",
            "Clear Templates Archive"});
            this.comboBox1.Location = new System.Drawing.Point(12, 31);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(168, 21);
            this.comboBox1.TabIndex = 5;
            this.comboBox1.SelectedIndexChanged += new System.EventHandler(this.comboBox1_SelectedIndexChanged);
            // 
            // readmeButton
            // 
            this.readmeButton.Location = new System.Drawing.Point(201, 15);
            this.readmeButton.Name = "readmeButton";
            this.readmeButton.Size = new System.Drawing.Size(177, 23);
            this.readmeButton.TabIndex = 0;
            this.readmeButton.Text = "Open README File";
            this.readmeButton.UseVisualStyleBackColor = true;
            this.readmeButton.Click += new System.EventHandler(this.button2_Click);
            // 
            // runExeButton
            // 
            this.runExeButton.Enabled = false;
            this.runExeButton.Location = new System.Drawing.Point(12, 272);
            this.runExeButton.Name = "runExeButton";
            this.runExeButton.Size = new System.Drawing.Size(168, 77);
            this.runExeButton.TabIndex = 12;
            this.runExeButton.Text = "Run Executable";
            this.runExeButton.UseVisualStyleBackColor = true;
            this.runExeButton.Click += new System.EventHandler(this.button1_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 15);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(131, 13);
            this.label1.TabIndex = 4;
            this.label1.Text = "Select Executable to Run:";
            // 
            // cleanSlateButton
            // 
            this.cleanSlateButton.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.cleanSlateButton.Cursor = System.Windows.Forms.Cursors.Arrow;
            this.cleanSlateButton.Enabled = false;
            this.cleanSlateButton.FlatAppearance.BorderColor = System.Drawing.SystemColors.ControlLight;
            this.cleanSlateButton.FlatAppearance.MouseDownBackColor = System.Drawing.SystemColors.ControlLight;
            this.cleanSlateButton.FlatAppearance.MouseOverBackColor = System.Drawing.SystemColors.ControlLight;
            this.cleanSlateButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.cleanSlateButton.Location = new System.Drawing.Point(-3, 493);
            this.cleanSlateButton.Name = "cleanSlateButton";
            this.cleanSlateButton.Size = new System.Drawing.Size(25, 26);
            this.cleanSlateButton.TabIndex = 17;
            this.cleanSlateButton.TabStop = false;
            this.cleanSlateButton.UseVisualStyleBackColor = true;
            this.cleanSlateButton.Click += new System.EventHandler(this.button3_Click);
            // 
            // inputInfoButton
            // 
            this.inputInfoButton.Location = new System.Drawing.Point(20, 20);
            this.inputInfoButton.Name = "inputInfoButton";
            this.inputInfoButton.Size = new System.Drawing.Size(136, 55);
            this.inputInfoButton.TabIndex = 0;
            this.inputInfoButton.Text = "Open Input Information Text File";
            this.inputInfoButton.UseVisualStyleBackColor = true;
            this.inputInfoButton.Click += new System.EventHandler(this.button4_Click);
            // 
            // userInputGroup
            // 
            this.userInputGroup.BackColor = System.Drawing.SystemColors.Control;
            this.userInputGroup.Controls.Add(this.patientInfoButton);
            this.userInputGroup.Controls.Add(this.conditionalButton);
            this.userInputGroup.Controls.Add(this.inputInfoButton);
            this.userInputGroup.Enabled = false;
            this.userInputGroup.Location = new System.Drawing.Point(201, 123);
            this.userInputGroup.Name = "userInputGroup";
            this.userInputGroup.Size = new System.Drawing.Size(177, 212);
            this.userInputGroup.TabIndex = 11;
            this.userInputGroup.TabStop = false;
            this.userInputGroup.Text = "User Input Locations";
            this.userInputGroup.Enter += new System.EventHandler(this.groupBox1_Enter);
            // 
            // patientInfoButton
            // 
            this.patientInfoButton.Location = new System.Drawing.Point(20, 81);
            this.patientInfoButton.Name = "patientInfoButton";
            this.patientInfoButton.Size = new System.Drawing.Size(136, 55);
            this.patientInfoButton.TabIndex = 1;
            this.patientInfoButton.Text = "Open SQL Extracted Information";
            this.patientInfoButton.UseVisualStyleBackColor = true;
            this.patientInfoButton.Click += new System.EventHandler(this.button7_Click);
            // 
            // conditionalButton
            // 
            this.conditionalButton.Location = new System.Drawing.Point(20, 142);
            this.conditionalButton.Name = "conditionalButton";
            this.conditionalButton.Size = new System.Drawing.Size(136, 55);
            this.conditionalButton.TabIndex = 2;
            this.conditionalButton.Text = "Open Conditional Template";
            this.conditionalButton.UseVisualStyleBackColor = true;
            this.conditionalButton.Click += new System.EventHandler(this.button5_Click);
            // 
            // archiveButton
            // 
            this.archiveButton.Enabled = false;
            this.archiveButton.Location = new System.Drawing.Point(13, 355);
            this.archiveButton.Name = "archiveButton";
            this.archiveButton.Size = new System.Drawing.Size(140, 31);
            this.archiveButton.TabIndex = 14;
            this.archiveButton.Text = "Open Template Archive";
            this.archiveButton.UseVisualStyleBackColor = true;
            this.archiveButton.Click += new System.EventHandler(this.button6_Click);
            // 
            // editingGroup
            // 
            this.editingGroup.BackColor = System.Drawing.SystemColors.Control;
            this.editingGroup.Controls.Add(this.singleTemplateButton);
            this.editingGroup.Controls.Add(this.templatesButton);
            this.editingGroup.Controls.Add(this.idListButton);
            this.editingGroup.Enabled = false;
            this.editingGroup.Location = new System.Drawing.Point(201, 352);
            this.editingGroup.Name = "editingGroup";
            this.editingGroup.Size = new System.Drawing.Size(177, 150);
            this.editingGroup.TabIndex = 13;
            this.editingGroup.TabStop = false;
            this.editingGroup.Text = "Editing Locations";
            this.editingGroup.Enter += new System.EventHandler(this.editingGroup_Enter);
            // 
            // singleTemplateButton
            // 
            this.singleTemplateButton.Enabled = false;
            this.singleTemplateButton.Location = new System.Drawing.Point(71, 19);
            this.singleTemplateButton.Name = "singleTemplateButton";
            this.singleTemplateButton.Size = new System.Drawing.Size(136, 55);
            this.singleTemplateButton.TabIndex = 1;
            this.singleTemplateButton.Text = "Open Patient Template";
            this.singleTemplateButton.UseVisualStyleBackColor = true;
            this.singleTemplateButton.Visible = false;
            this.singleTemplateButton.Click += new System.EventHandler(this.singleTemplateButton_Click);
            // 
            // templatesButton
            // 
            this.templatesButton.Location = new System.Drawing.Point(20, 19);
            this.templatesButton.Name = "templatesButton";
            this.templatesButton.Size = new System.Drawing.Size(136, 55);
            this.templatesButton.TabIndex = 0;
            this.templatesButton.Text = "Open Patient Template Location";
            this.templatesButton.UseVisualStyleBackColor = true;
            this.templatesButton.Click += new System.EventHandler(this.button9_Click);
            // 
            // idListButton
            // 
            this.idListButton.Location = new System.Drawing.Point(20, 80);
            this.idListButton.Name = "idListButton";
            this.idListButton.Size = new System.Drawing.Size(136, 55);
            this.idListButton.TabIndex = 2;
            this.idListButton.Text = "Open Patient ID List";
            this.idListButton.UseVisualStyleBackColor = true;
            this.idListButton.Click += new System.EventHandler(this.button8_Click);
            // 
            // failedButton
            // 
            this.failedButton.Enabled = false;
            this.failedButton.Location = new System.Drawing.Point(13, 392);
            this.failedButton.Name = "failedButton";
            this.failedButton.Size = new System.Drawing.Size(140, 30);
            this.failedButton.TabIndex = 15;
            this.failedButton.Text = "Open Failed Extractions";
            this.failedButton.UseVisualStyleBackColor = true;
            this.failedButton.Click += new System.EventHandler(this.button10_Click);
            // 
            // spreadsheetButton
            // 
            this.spreadsheetButton.Enabled = false;
            this.spreadsheetButton.Location = new System.Drawing.Point(13, 428);
            this.spreadsheetButton.Name = "spreadsheetButton";
            this.spreadsheetButton.Size = new System.Drawing.Size(140, 30);
            this.spreadsheetButton.TabIndex = 16;
            this.spreadsheetButton.Text = "Open Spreadsheet";
            this.spreadsheetButton.UseVisualStyleBackColor = true;
            this.spreadsheetButton.Click += new System.EventHandler(this.button11_Click);
            // 
            // overrideButton
            // 
            this.overrideButton.Location = new System.Drawing.Point(201, 65);
            this.overrideButton.Name = "overrideButton";
            this.overrideButton.Size = new System.Drawing.Size(177, 23);
            this.overrideButton.TabIndex = 2;
            this.overrideButton.Text = "Enable All Editing Buttons";
            this.overrideButton.UseVisualStyleBackColor = true;
            this.overrideButton.Click += new System.EventHandler(this.button1_Click_1);
            // 
            // groupBoxStructure
            // 
            this.groupBoxStructure.Controls.Add(this.linkLabelDesAll);
            this.groupBoxStructure.Controls.Add(this.linkLabelSelAll);
            this.groupBoxStructure.Controls.Add(this.groupBoxPlanSum);
            this.groupBoxStructure.Controls.Add(this.checkBoxInput);
            this.groupBoxStructure.Controls.Add(this.checkBoxPros);
            this.groupBoxStructure.Controls.Add(this.checkBoxLung);
            this.groupBoxStructure.Controls.Add(this.checkBoxBrea);
            this.groupBoxStructure.Controls.Add(this.checkBoxBrai);
            this.groupBoxStructure.Controls.Add(this.textBoxInput);
            this.groupBoxStructure.Enabled = false;
            this.groupBoxStructure.Location = new System.Drawing.Point(12, 99);
            this.groupBoxStructure.Name = "groupBoxStructure";
            this.groupBoxStructure.Size = new System.Drawing.Size(168, 160);
            this.groupBoxStructure.TabIndex = 10;
            this.groupBoxStructure.TabStop = false;
            this.groupBoxStructure.Text = "Select Site to Run On:";
            this.groupBoxStructure.Enter += new System.EventHandler(this.groupBox2_Enter);
            // 
            // linkLabelDesAll
            // 
            this.linkLabelDesAll.AutoSize = true;
            this.linkLabelDesAll.Enabled = false;
            this.linkLabelDesAll.Location = new System.Drawing.Point(13, 16);
            this.linkLabelDesAll.Name = "linkLabelDesAll";
            this.linkLabelDesAll.Size = new System.Drawing.Size(63, 13);
            this.linkLabelDesAll.TabIndex = 0;
            this.linkLabelDesAll.TabStop = true;
            this.linkLabelDesAll.Text = "Deselect All";
            this.linkLabelDesAll.Visible = false;
            this.linkLabelDesAll.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabelDesAll_LinkClicked);
            // 
            // linkLabelSelAll
            // 
            this.linkLabelSelAll.AutoSize = true;
            this.linkLabelSelAll.Location = new System.Drawing.Point(13, 16);
            this.linkLabelSelAll.Name = "linkLabelSelAll";
            this.linkLabelSelAll.Size = new System.Drawing.Size(51, 13);
            this.linkLabelSelAll.TabIndex = 1;
            this.linkLabelSelAll.TabStop = true;
            this.linkLabelSelAll.Text = "Select All";
            this.linkLabelSelAll.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabelSelAll_LinkClicked);
            // 
            // groupBoxPlanSum
            // 
            this.groupBoxPlanSum.BackColor = System.Drawing.SystemColors.ControlLight;
            this.groupBoxPlanSum.Controls.Add(this.label3);
            this.groupBoxPlanSum.Controls.Add(this.checkPlanSumBrai);
            this.groupBoxPlanSum.Controls.Add(this.checkPlanSumLung);
            this.groupBoxPlanSum.Controls.Add(this.checkPlanSumBrea);
            this.groupBoxPlanSum.Controls.Add(this.checkPlanSumInput);
            this.groupBoxPlanSum.Controls.Add(this.checkPlanSumPros);
            this.groupBoxPlanSum.Enabled = false;
            this.groupBoxPlanSum.Location = new System.Drawing.Point(101, 19);
            this.groupBoxPlanSum.Name = "groupBoxPlanSum";
            this.groupBoxPlanSum.Size = new System.Drawing.Size(61, 135);
            this.groupBoxPlanSum.TabIndex = 8;
            this.groupBoxPlanSum.TabStop = false;
            this.groupBoxPlanSum.Visible = false;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(1, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(55, 13);
            this.label3.TabIndex = 1;
            this.label3.Text = "Plan Sum:";
            // 
            // checkPlanSumBrai
            // 
            this.checkPlanSumBrai.AutoSize = true;
            this.checkPlanSumBrai.Enabled = false;
            this.checkPlanSumBrai.Location = new System.Drawing.Point(23, 16);
            this.checkPlanSumBrai.Name = "checkPlanSumBrai";
            this.checkPlanSumBrai.Size = new System.Drawing.Size(15, 14);
            this.checkPlanSumBrai.TabIndex = 2;
            this.toolTipBrai.SetToolTip(this.checkPlanSumBrai, "Plan Sum not available for Brain");
            this.checkPlanSumBrai.UseVisualStyleBackColor = true;
            this.checkPlanSumBrai.CheckedChanged += new System.EventHandler(this.checkPlanSumBrai_CheckedChanged);
            // 
            // checkPlanSumLung
            // 
            this.checkPlanSumLung.AutoSize = true;
            this.checkPlanSumLung.Enabled = false;
            this.checkPlanSumLung.Location = new System.Drawing.Point(23, 62);
            this.checkPlanSumLung.Name = "checkPlanSumLung";
            this.checkPlanSumLung.Size = new System.Drawing.Size(15, 14);
            this.checkPlanSumLung.TabIndex = 4;
            this.toolTipLung.SetToolTip(this.checkPlanSumLung, "Plan Sum not available for Lung");
            this.checkPlanSumLung.UseVisualStyleBackColor = true;
            // 
            // checkPlanSumBrea
            // 
            this.checkPlanSumBrea.AutoSize = true;
            this.checkPlanSumBrea.Enabled = false;
            this.checkPlanSumBrea.Location = new System.Drawing.Point(23, 39);
            this.checkPlanSumBrea.Name = "checkPlanSumBrea";
            this.checkPlanSumBrea.Size = new System.Drawing.Size(15, 14);
            this.checkPlanSumBrea.TabIndex = 3;
            this.toolTipBrea.SetToolTip(this.checkPlanSumBrea, "Plan Sum not available for Breast");
            this.checkPlanSumBrea.UseVisualStyleBackColor = true;
            // 
            // checkPlanSumInput
            // 
            this.checkPlanSumInput.AutoSize = true;
            this.checkPlanSumInput.Enabled = false;
            this.checkPlanSumInput.Location = new System.Drawing.Point(23, 109);
            this.checkPlanSumInput.Name = "checkPlanSumInput";
            this.checkPlanSumInput.Size = new System.Drawing.Size(15, 14);
            this.checkPlanSumInput.TabIndex = 0;
            this.checkPlanSumInput.UseVisualStyleBackColor = true;
            this.checkPlanSumInput.CheckedChanged += new System.EventHandler(this.checkPlanSumInput_CheckedChanged);
            // 
            // checkPlanSumPros
            // 
            this.checkPlanSumPros.AutoSize = true;
            this.checkPlanSumPros.Enabled = false;
            this.checkPlanSumPros.Location = new System.Drawing.Point(23, 85);
            this.checkPlanSumPros.Name = "checkPlanSumPros";
            this.checkPlanSumPros.Size = new System.Drawing.Size(15, 14);
            this.checkPlanSumPros.TabIndex = 5;
            this.checkPlanSumPros.UseVisualStyleBackColor = true;
            this.checkPlanSumPros.CheckedChanged += new System.EventHandler(this.checkPlanSumPros_CheckedChanged);
            // 
            // checkBoxInput
            // 
            this.checkBoxInput.AutoSize = true;
            this.checkBoxInput.Location = new System.Drawing.Point(14, 128);
            this.checkBoxInput.Name = "checkBoxInput";
            this.checkBoxInput.Size = new System.Drawing.Size(15, 14);
            this.checkBoxInput.TabIndex = 6;
            this.checkBoxInput.UseVisualStyleBackColor = true;
            this.checkBoxInput.CheckedChanged += new System.EventHandler(this.checkBoxInput_CheckedChanged);
            // 
            // checkBoxPros
            // 
            this.checkBoxPros.AutoSize = true;
            this.checkBoxPros.Location = new System.Drawing.Point(14, 103);
            this.checkBoxPros.Name = "checkBoxPros";
            this.checkBoxPros.Size = new System.Drawing.Size(65, 17);
            this.checkBoxPros.TabIndex = 5;
            this.checkBoxPros.Text = "Prostate";
            this.checkBoxPros.UseVisualStyleBackColor = true;
            this.checkBoxPros.CheckedChanged += new System.EventHandler(this.checkBoxPros_CheckedChanged);
            // 
            // checkBoxLung
            // 
            this.checkBoxLung.AutoSize = true;
            this.checkBoxLung.Location = new System.Drawing.Point(14, 80);
            this.checkBoxLung.Name = "checkBoxLung";
            this.checkBoxLung.Size = new System.Drawing.Size(50, 17);
            this.checkBoxLung.TabIndex = 4;
            this.checkBoxLung.Text = "Lung";
            this.checkBoxLung.UseVisualStyleBackColor = true;
            this.checkBoxLung.CheckedChanged += new System.EventHandler(this.checkBoxLung_CheckedChanged);
            // 
            // checkBoxBrea
            // 
            this.checkBoxBrea.AutoSize = true;
            this.checkBoxBrea.Location = new System.Drawing.Point(14, 57);
            this.checkBoxBrea.Name = "checkBoxBrea";
            this.checkBoxBrea.Size = new System.Drawing.Size(56, 17);
            this.checkBoxBrea.TabIndex = 3;
            this.checkBoxBrea.Text = "Breast";
            this.checkBoxBrea.UseVisualStyleBackColor = true;
            this.checkBoxBrea.CheckedChanged += new System.EventHandler(this.checkBoxBrea_CheckedChanged);
            // 
            // checkBoxBrai
            // 
            this.checkBoxBrai.AutoSize = true;
            this.checkBoxBrai.Location = new System.Drawing.Point(14, 34);
            this.checkBoxBrai.Name = "checkBoxBrai";
            this.checkBoxBrai.Size = new System.Drawing.Size(50, 17);
            this.checkBoxBrai.TabIndex = 2;
            this.checkBoxBrai.Text = "Brain";
            this.checkBoxBrai.UseVisualStyleBackColor = true;
            this.checkBoxBrai.CheckedChanged += new System.EventHandler(this.checkBoxBrai_CheckedChanged);
            // 
            // textBoxInput
            // 
            this.textBoxInput.Location = new System.Drawing.Point(32, 125);
            this.textBoxInput.MaxLength = 12;
            this.textBoxInput.Name = "textBoxInput";
            this.textBoxInput.Size = new System.Drawing.Size(63, 20);
            this.textBoxInput.TabIndex = 7;
            this.textBoxInput.TabStop = false;
            this.textBoxInput.TextChanged += new System.EventHandler(this.textBox1_TextChanged_2);
            // 
            // feedbackButton
            // 
            this.feedbackButton.Location = new System.Drawing.Point(201, 90);
            this.feedbackButton.Name = "feedbackButton";
            this.feedbackButton.Size = new System.Drawing.Size(177, 23);
            this.feedbackButton.TabIndex = 3;
            this.feedbackButton.Text = "Give Feedback";
            this.feedbackButton.UseVisualStyleBackColor = true;
            this.feedbackButton.Click += new System.EventHandler(this.feedbackButton_Click);
            // 
            // textBoxID
            // 
            this.textBoxID.Enabled = false;
            this.textBoxID.Location = new System.Drawing.Point(13, 75);
            this.textBoxID.MaxLength = 8;
            this.textBoxID.Name = "textBoxID";
            this.textBoxID.Size = new System.Drawing.Size(166, 20);
            this.textBoxID.TabIndex = 9;
            this.textBoxID.TextChanged += new System.EventHandler(this.textBoxID_TextChanged);
            // 
            // labelID
            // 
            this.labelID.AutoSize = true;
            this.labelID.BackColor = System.Drawing.SystemColors.Window;
            this.labelID.Enabled = false;
            this.labelID.Location = new System.Drawing.Point(17, 78);
            this.labelID.Name = "labelID";
            this.labelID.Size = new System.Drawing.Size(91, 13);
            this.labelID.TabIndex = 8;
            this.labelID.Text = "Patient ID Search";
            this.labelID.Click += new System.EventHandler(this.labelID_Click);
            // 
            // buttonQuery
            // 
            this.buttonQuery.Location = new System.Drawing.Point(201, 40);
            this.buttonQuery.Name = "buttonQuery";
            this.buttonQuery.Size = new System.Drawing.Size(177, 23);
            this.buttonQuery.TabIndex = 1;
            this.buttonQuery.Text = "Database Query";
            this.buttonQuery.UseVisualStyleBackColor = true;
            this.buttonQuery.Click += new System.EventHandler(this.buttonQuery_Click);
            // 
            // checkBoxDose
            // 
            this.checkBoxDose.AutoSize = true;
            this.checkBoxDose.Enabled = false;
            this.checkBoxDose.Location = new System.Drawing.Point(18, 55);
            this.checkBoxDose.Name = "checkBoxDose";
            this.checkBoxDose.Size = new System.Drawing.Size(51, 17);
            this.checkBoxDose.TabIndex = 6;
            this.checkBoxDose.Text = "Dose";
            this.checkBoxDose.UseVisualStyleBackColor = true;
            this.checkBoxDose.CheckedChanged += new System.EventHandler(this.checkBoxDose_CheckedChanged);
            // 
            // checkBoxDVH
            // 
            this.checkBoxDVH.AutoSize = true;
            this.checkBoxDVH.Enabled = false;
            this.checkBoxDVH.Location = new System.Drawing.Point(75, 55);
            this.checkBoxDVH.Name = "checkBoxDVH";
            this.checkBoxDVH.Size = new System.Drawing.Size(49, 17);
            this.checkBoxDVH.TabIndex = 7;
            this.checkBoxDVH.Text = "DVH";
            this.checkBoxDVH.UseVisualStyleBackColor = true;
            this.checkBoxDVH.CheckedChanged += new System.EventHandler(this.checkBoxDVH_CheckedChanged);
            // 
            // toolTipBrai
            // 
            this.toolTipBrai.Popup += new System.Windows.Forms.PopupEventHandler(this.toolTipBrai_Popup);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlLight;
            this.ClientSize = new System.Drawing.Size(401, 517);
            this.Controls.Add(this.checkBoxDVH);
            this.Controls.Add(this.checkBoxDose);
            this.Controls.Add(this.buttonQuery);
            this.Controls.Add(this.labelID);
            this.Controls.Add(this.textBoxID);
            this.Controls.Add(this.feedbackButton);
            this.Controls.Add(this.groupBoxStructure);
            this.Controls.Add(this.overrideButton);
            this.Controls.Add(this.spreadsheetButton);
            this.Controls.Add(this.failedButton);
            this.Controls.Add(this.editingGroup);
            this.Controls.Add(this.archiveButton);
            this.Controls.Add(this.userInputGroup);
            this.Controls.Add(this.cleanSlateButton);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.runExeButton);
            this.Controls.Add(this.readmeButton);
            this.Controls.Add(this.comboBox1);
            this.Cursor = System.Windows.Forms.Cursors.Arrow;
            this.MaximumSize = new System.Drawing.Size(417, 555);
            this.MinimumSize = new System.Drawing.Size(417, 555);
            this.Name = "Form1";
            this.Text = "Extract Dose/DVH";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.userInputGroup.ResumeLayout(false);
            this.editingGroup.ResumeLayout(false);
            this.groupBoxStructure.ResumeLayout(false);
            this.groupBoxStructure.PerformLayout();
            this.groupBoxPlanSum.ResumeLayout(false);
            this.groupBoxPlanSum.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.Button readmeButton;
        private System.Windows.Forms.Button runExeButton;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button cleanSlateButton;
        private System.Windows.Forms.Button inputInfoButton;
        private System.Windows.Forms.GroupBox userInputGroup;
        private System.Windows.Forms.Button conditionalButton;
        private System.Windows.Forms.Button archiveButton;
        private System.Windows.Forms.GroupBox editingGroup;
        private System.Windows.Forms.Button patientInfoButton;
        private System.Windows.Forms.Button templatesButton;
        private System.Windows.Forms.Button idListButton;
        private System.Windows.Forms.Button failedButton;
        private System.Windows.Forms.Button spreadsheetButton;
        private System.Windows.Forms.Button overrideButton;
        private System.Windows.Forms.GroupBox groupBoxStructure;
        private System.Windows.Forms.TextBox textBoxInput;
        private System.Windows.Forms.Button feedbackButton;
        private System.Windows.Forms.TextBox textBoxID;
        private System.Windows.Forms.Label labelID;
        private System.Windows.Forms.Button singleTemplateButton;
		private System.Windows.Forms.Button buttonQuery;
		private System.Windows.Forms.CheckBox checkBoxInput;
		private System.Windows.Forms.CheckBox checkBoxPros;
		private System.Windows.Forms.CheckBox checkBoxLung;
		private System.Windows.Forms.CheckBox checkBoxBrea;
		private System.Windows.Forms.CheckBox checkBoxBrai;
		private System.Windows.Forms.CheckBox checkBoxDose;
		private System.Windows.Forms.CheckBox checkBoxDVH;
		private System.Windows.Forms.GroupBox groupBoxPlanSum;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.CheckBox checkPlanSumBrai;
		private System.Windows.Forms.CheckBox checkPlanSumLung;
		private System.Windows.Forms.CheckBox checkPlanSumBrea;
		private System.Windows.Forms.CheckBox checkPlanSumInput;
		private System.Windows.Forms.CheckBox checkPlanSumPros;
		private System.Windows.Forms.ToolTip toolTipBrai;
		private System.Windows.Forms.LinkLabel linkLabelSelAll;
		private System.Windows.Forms.LinkLabel linkLabelDesAll;
        private System.Windows.Forms.ToolTip toolTipLung;
        private System.Windows.Forms.ToolTip toolTipBrea;
    }
}

