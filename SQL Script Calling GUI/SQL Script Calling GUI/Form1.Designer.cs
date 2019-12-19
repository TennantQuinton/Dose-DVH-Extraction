namespace SQL_Script_Calling_GUI
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
            this.buttonQuery = new System.Windows.Forms.Button();
            this.groupBoxDatabase = new System.Windows.Forms.GroupBox();
            this.textBoxPassword = new System.Windows.Forms.TextBox();
            this.textBoxUserID = new System.Windows.Forms.TextBox();
            this.textBoxDatabase = new System.Windows.Forms.TextBox();
            this.labelPassword = new System.Windows.Forms.Label();
            this.labelUserID = new System.Windows.Forms.Label();
            this.labelDatabase = new System.Windows.Forms.Label();
            this.labelServer = new System.Windows.Forms.Label();
            this.textBoxServer = new System.Windows.Forms.TextBox();
            this.groupBoxQuery = new System.Windows.Forms.GroupBox();
            this.textBoxVolCode = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.dateTimePickerEnd = new System.Windows.Forms.DateTimePicker();
            this.dateTimePickerStart = new System.Windows.Forms.DateTimePicker();
            this.comboBoxSite = new System.Windows.Forms.ComboBox();
            this.labelStructure = new System.Windows.Forms.Label();
            this.labelEnd = new System.Windows.Forms.Label();
            this.labelStart = new System.Windows.Forms.Label();
            this.textBoxExcel = new System.Windows.Forms.TextBox();
            this.labelExcel = new System.Windows.Forms.Label();
            this.buttonExtract = new System.Windows.Forms.Button();
            this.linkLabelAdvQ = new System.Windows.Forms.LinkLabel();
            this.groupBoxAdvQ = new System.Windows.Forms.GroupBox();
            this.textPatIDCond = new System.Windows.Forms.TextBox();
            this.comboPatIDLike = new System.Windows.Forms.ComboBox();
            this.checkPatID = new System.Windows.Forms.CheckBox();
            this.comboMLCLike = new System.Windows.Forms.ComboBox();
            this.comboGantLike = new System.Windows.Forms.ComboBox();
            this.numericFract = new System.Windows.Forms.NumericUpDown();
            this.numericDose = new System.Windows.Forms.NumericUpDown();
            this.comboFractEq = new System.Windows.Forms.ComboBox();
            this.comboDoseEq = new System.Windows.Forms.ComboBox();
            this.comboMLCCon = new System.Windows.Forms.ComboBox();
            this.comboGantryCond = new System.Windows.Forms.ComboBox();
            this.checkNFract = new System.Windows.Forms.CheckBox();
            this.checkPresDose = new System.Windows.Forms.CheckBox();
            this.checkMLCPlan = new System.Windows.Forms.CheckBox();
            this.checkGantryRot = new System.Windows.Forms.CheckBox();
            this.linkLabelCancel = new System.Windows.Forms.LinkLabel();
            this.groupBoxDatabase.SuspendLayout();
            this.groupBoxQuery.SuspendLayout();
            this.groupBoxAdvQ.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numericFract)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericDose)).BeginInit();
            this.SuspendLayout();
            // 
            // buttonQuery
            // 
            this.buttonQuery.Enabled = false;
            this.buttonQuery.Location = new System.Drawing.Point(179, 313);
            this.buttonQuery.Name = "buttonQuery";
            this.buttonQuery.Size = new System.Drawing.Size(126, 29);
            this.buttonQuery.TabIndex = 7;
            this.buttonQuery.Text = "Query";
            this.buttonQuery.UseVisualStyleBackColor = true;
            this.buttonQuery.Click += new System.EventHandler(this.button1_Click);
            // 
            // groupBoxDatabase
            // 
            this.groupBoxDatabase.Controls.Add(this.textBoxPassword);
            this.groupBoxDatabase.Controls.Add(this.textBoxUserID);
            this.groupBoxDatabase.Controls.Add(this.textBoxDatabase);
            this.groupBoxDatabase.Controls.Add(this.labelPassword);
            this.groupBoxDatabase.Controls.Add(this.labelUserID);
            this.groupBoxDatabase.Controls.Add(this.labelDatabase);
            this.groupBoxDatabase.Controls.Add(this.labelServer);
            this.groupBoxDatabase.Controls.Add(this.textBoxServer);
            this.groupBoxDatabase.Location = new System.Drawing.Point(12, 30);
            this.groupBoxDatabase.Name = "groupBoxDatabase";
            this.groupBoxDatabase.Size = new System.Drawing.Size(293, 133);
            this.groupBoxDatabase.TabIndex = 1;
            this.groupBoxDatabase.TabStop = false;
            this.groupBoxDatabase.Text = "Database Info";
            this.groupBoxDatabase.Enter += new System.EventHandler(this.groupBoxDatabase_Enter);
            // 
            // textBoxPassword
            // 
            this.textBoxPassword.Enabled = false;
            this.textBoxPassword.Location = new System.Drawing.Point(93, 88);
            this.textBoxPassword.Name = "textBoxPassword";
            this.textBoxPassword.PasswordChar = '*';
            this.textBoxPassword.Size = new System.Drawing.Size(127, 20);
            this.textBoxPassword.TabIndex = 7;
            this.textBoxPassword.Text = "reports";
            this.textBoxPassword.TextChanged += new System.EventHandler(this.textBoxPassword_TextChanged);
            // 
            // textBoxUserID
            // 
            this.textBoxUserID.Enabled = false;
            this.textBoxUserID.Location = new System.Drawing.Point(93, 66);
            this.textBoxUserID.Name = "textBoxUserID";
            this.textBoxUserID.Size = new System.Drawing.Size(127, 20);
            this.textBoxUserID.TabIndex = 5;
            this.textBoxUserID.Text = "reports";
            this.textBoxUserID.TextChanged += new System.EventHandler(this.textBoxUserID_TextChanged);
            // 
            // textBoxDatabase
            // 
            this.textBoxDatabase.Enabled = false;
            this.textBoxDatabase.Location = new System.Drawing.Point(93, 44);
            this.textBoxDatabase.Name = "textBoxDatabase";
            this.textBoxDatabase.Size = new System.Drawing.Size(127, 20);
            this.textBoxDatabase.TabIndex = 3;
            this.textBoxDatabase.Text = "variansystem";
            this.textBoxDatabase.TextChanged += new System.EventHandler(this.textBoxDatabase_TextChanged);
            // 
            // labelPassword
            // 
            this.labelPassword.AutoSize = true;
            this.labelPassword.Location = new System.Drawing.Point(20, 91);
            this.labelPassword.Name = "labelPassword";
            this.labelPassword.Size = new System.Drawing.Size(56, 13);
            this.labelPassword.TabIndex = 6;
            this.labelPassword.Text = "Password:";
            // 
            // labelUserID
            // 
            this.labelUserID.AutoSize = true;
            this.labelUserID.Location = new System.Drawing.Point(20, 69);
            this.labelUserID.Name = "labelUserID";
            this.labelUserID.Size = new System.Drawing.Size(46, 13);
            this.labelUserID.TabIndex = 4;
            this.labelUserID.Text = "User ID:";
            // 
            // labelDatabase
            // 
            this.labelDatabase.AutoSize = true;
            this.labelDatabase.Location = new System.Drawing.Point(20, 47);
            this.labelDatabase.Name = "labelDatabase";
            this.labelDatabase.Size = new System.Drawing.Size(56, 13);
            this.labelDatabase.TabIndex = 2;
            this.labelDatabase.Text = "Database:";
            // 
            // labelServer
            // 
            this.labelServer.AutoSize = true;
            this.labelServer.Location = new System.Drawing.Point(20, 25);
            this.labelServer.Name = "labelServer";
            this.labelServer.Size = new System.Drawing.Size(41, 13);
            this.labelServer.TabIndex = 0;
            this.labelServer.Text = "Server:";
            // 
            // textBoxServer
            // 
            this.textBoxServer.Location = new System.Drawing.Point(93, 22);
            this.textBoxServer.Name = "textBoxServer";
            this.textBoxServer.Size = new System.Drawing.Size(127, 20);
            this.textBoxServer.TabIndex = 1;
            this.textBoxServer.Text = "grc505n";
            this.textBoxServer.TextChanged += new System.EventHandler(this.textBoxServer_TextChanged);
            // 
            // groupBoxQuery
            // 
            this.groupBoxQuery.Controls.Add(this.textBoxVolCode);
            this.groupBoxQuery.Controls.Add(this.label1);
            this.groupBoxQuery.Controls.Add(this.dateTimePickerEnd);
            this.groupBoxQuery.Controls.Add(this.dateTimePickerStart);
            this.groupBoxQuery.Controls.Add(this.comboBoxSite);
            this.groupBoxQuery.Controls.Add(this.labelStructure);
            this.groupBoxQuery.Controls.Add(this.labelEnd);
            this.groupBoxQuery.Controls.Add(this.labelStart);
            this.groupBoxQuery.Enabled = false;
            this.groupBoxQuery.Location = new System.Drawing.Point(12, 169);
            this.groupBoxQuery.Name = "groupBoxQuery";
            this.groupBoxQuery.Size = new System.Drawing.Size(293, 119);
            this.groupBoxQuery.TabIndex = 2;
            this.groupBoxQuery.TabStop = false;
            this.groupBoxQuery.Text = "Query";
            this.groupBoxQuery.Enter += new System.EventHandler(this.groupBoxQuery_Enter);
            // 
            // textBoxVolCode
            // 
            this.textBoxVolCode.Enabled = false;
            this.textBoxVolCode.Location = new System.Drawing.Point(93, 90);
            this.textBoxVolCode.Name = "textBoxVolCode";
            this.textBoxVolCode.Size = new System.Drawing.Size(127, 20);
            this.textBoxVolCode.TabIndex = 7;
            this.textBoxVolCode.TextChanged += new System.EventHandler(this.textBoxVolCode_TextChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(20, 94);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(73, 13);
            this.label1.TabIndex = 6;
            this.label1.Text = "Volume Code:";
            // 
            // dateTimePickerEnd
            // 
            this.dateTimePickerEnd.Enabled = false;
            this.dateTimePickerEnd.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dateTimePickerEnd.Location = new System.Drawing.Point(93, 42);
            this.dateTimePickerEnd.MaxDate = new System.DateTime(3003, 12, 31, 0, 0, 0, 0);
            this.dateTimePickerEnd.MinDate = new System.DateTime(2003, 1, 1, 0, 0, 0, 0);
            this.dateTimePickerEnd.Name = "dateTimePickerEnd";
            this.dateTimePickerEnd.Size = new System.Drawing.Size(127, 20);
            this.dateTimePickerEnd.TabIndex = 3;
            this.dateTimePickerEnd.Value = new System.DateTime(2003, 1, 1, 0, 0, 0, 0);
            this.dateTimePickerEnd.ValueChanged += new System.EventHandler(this.dateTimePickerEnd_ValueChanged);
            // 
            // dateTimePickerStart
            // 
            this.dateTimePickerStart.Enabled = false;
            this.dateTimePickerStart.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dateTimePickerStart.Location = new System.Drawing.Point(93, 20);
            this.dateTimePickerStart.MaxDate = new System.DateTime(3003, 12, 31, 0, 0, 0, 0);
            this.dateTimePickerStart.MinDate = new System.DateTime(2003, 1, 1, 0, 0, 0, 0);
            this.dateTimePickerStart.Name = "dateTimePickerStart";
            this.dateTimePickerStart.Size = new System.Drawing.Size(127, 20);
            this.dateTimePickerStart.TabIndex = 1;
            this.dateTimePickerStart.Value = new System.DateTime(2003, 1, 1, 0, 0, 0, 0);
            this.dateTimePickerStart.ValueChanged += new System.EventHandler(this.dateTimePickerStart_ValueChanged);
            // 
            // comboBoxSite
            // 
            this.comboBoxSite.Enabled = false;
            this.comboBoxSite.FormattingEnabled = true;
            this.comboBoxSite.Items.AddRange(new object[] {
            " ",
            "Brain",
            "Breast",
            "Lung",
            "Prostate",
            "All Patients"});
            this.comboBoxSite.Location = new System.Drawing.Point(93, 67);
            this.comboBoxSite.Name = "comboBoxSite";
            this.comboBoxSite.Size = new System.Drawing.Size(127, 21);
            this.comboBoxSite.TabIndex = 5;
            this.comboBoxSite.SelectedIndexChanged += new System.EventHandler(this.comboBoxStructure_SelectedIndexChanged);
            // 
            // labelStructure
            // 
            this.labelStructure.AutoSize = true;
            this.labelStructure.Location = new System.Drawing.Point(20, 70);
            this.labelStructure.Name = "labelStructure";
            this.labelStructure.Size = new System.Drawing.Size(69, 13);
            this.labelStructure.TabIndex = 4;
            this.labelStructure.Text = "Disease Site:";
            // 
            // labelEnd
            // 
            this.labelEnd.AutoSize = true;
            this.labelEnd.Location = new System.Drawing.Point(20, 48);
            this.labelEnd.Name = "labelEnd";
            this.labelEnd.Size = new System.Drawing.Size(55, 13);
            this.labelEnd.TabIndex = 2;
            this.labelEnd.Text = "End Date:";
            // 
            // labelStart
            // 
            this.labelStart.AutoSize = true;
            this.labelStart.Location = new System.Drawing.Point(20, 26);
            this.labelStart.Name = "labelStart";
            this.labelStart.Size = new System.Drawing.Size(58, 13);
            this.labelStart.TabIndex = 0;
            this.labelStart.Text = "Start Date:";
            // 
            // textBoxExcel
            // 
            this.textBoxExcel.Enabled = false;
            this.textBoxExcel.Location = new System.Drawing.Point(82, 291);
            this.textBoxExcel.Name = "textBoxExcel";
            this.textBoxExcel.Size = new System.Drawing.Size(223, 20);
            this.textBoxExcel.TabIndex = 4;
            this.textBoxExcel.TextChanged += new System.EventHandler(this.textBoxExcel_TextChanged);
            // 
            // labelExcel
            // 
            this.labelExcel.AutoSize = true;
            this.labelExcel.Location = new System.Drawing.Point(9, 294);
            this.labelExcel.Name = "labelExcel";
            this.labelExcel.Size = new System.Drawing.Size(67, 13);
            this.labelExcel.TabIndex = 3;
            this.labelExcel.Text = "Excel Name:";
            this.labelExcel.Click += new System.EventHandler(this.labelExcel_Click);
            // 
            // buttonExtract
            // 
            this.buttonExtract.Location = new System.Drawing.Point(179, 9);
            this.buttonExtract.Name = "buttonExtract";
            this.buttonExtract.Size = new System.Drawing.Size(126, 21);
            this.buttonExtract.TabIndex = 0;
            this.buttonExtract.Text = "Extract Dose/DVH";
            this.buttonExtract.UseVisualStyleBackColor = true;
            this.buttonExtract.Click += new System.EventHandler(this.button1_Click_1);
            // 
            // linkLabelAdvQ
            // 
            this.linkLabelAdvQ.AutoSize = true;
            this.linkLabelAdvQ.Location = new System.Drawing.Point(86, 314);
            this.linkLabelAdvQ.Name = "linkLabelAdvQ";
            this.linkLabelAdvQ.Size = new System.Drawing.Size(87, 13);
            this.linkLabelAdvQ.TabIndex = 6;
            this.linkLabelAdvQ.TabStop = true;
            this.linkLabelAdvQ.Text = "Advanced Query";
            this.linkLabelAdvQ.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel1_LinkClicked);
            // 
            // groupBoxAdvQ
            // 
            this.groupBoxAdvQ.Controls.Add(this.textPatIDCond);
            this.groupBoxAdvQ.Controls.Add(this.comboPatIDLike);
            this.groupBoxAdvQ.Controls.Add(this.checkPatID);
            this.groupBoxAdvQ.Controls.Add(this.comboMLCLike);
            this.groupBoxAdvQ.Controls.Add(this.comboGantLike);
            this.groupBoxAdvQ.Controls.Add(this.numericFract);
            this.groupBoxAdvQ.Controls.Add(this.numericDose);
            this.groupBoxAdvQ.Controls.Add(this.comboFractEq);
            this.groupBoxAdvQ.Controls.Add(this.comboDoseEq);
            this.groupBoxAdvQ.Controls.Add(this.comboMLCCon);
            this.groupBoxAdvQ.Controls.Add(this.comboGantryCond);
            this.groupBoxAdvQ.Controls.Add(this.checkNFract);
            this.groupBoxAdvQ.Controls.Add(this.checkPresDose);
            this.groupBoxAdvQ.Controls.Add(this.checkMLCPlan);
            this.groupBoxAdvQ.Controls.Add(this.checkGantryRot);
            this.groupBoxAdvQ.Location = new System.Drawing.Point(258, 359);
            this.groupBoxAdvQ.Name = "groupBoxAdvQ";
            this.groupBoxAdvQ.Size = new System.Drawing.Size(293, 151);
            this.groupBoxAdvQ.TabIndex = 5;
            this.groupBoxAdvQ.TabStop = false;
            this.groupBoxAdvQ.Text = "Advanced Query";
            this.groupBoxAdvQ.Visible = false;
            // 
            // textPatIDCond
            // 
            this.textPatIDCond.Location = new System.Drawing.Point(206, 123);
            this.textPatIDCond.Name = "textPatIDCond";
            this.textPatIDCond.Size = new System.Drawing.Size(77, 20);
            this.textPatIDCond.TabIndex = 14;
            // 
            // comboPatIDLike
            // 
            this.comboPatIDLike.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboPatIDLike.FormattingEnabled = true;
            this.comboPatIDLike.Items.AddRange(new object[] {
            "like",
            "not like"});
            this.comboPatIDLike.Location = new System.Drawing.Point(123, 121);
            this.comboPatIDLike.Name = "comboPatIDLike";
            this.comboPatIDLike.Size = new System.Drawing.Size(75, 21);
            this.comboPatIDLike.TabIndex = 13;
            // 
            // checkPatID
            // 
            this.checkPatID.AutoSize = true;
            this.checkPatID.Location = new System.Drawing.Point(17, 121);
            this.checkPatID.Name = "checkPatID";
            this.checkPatID.Size = new System.Drawing.Size(71, 17);
            this.checkPatID.TabIndex = 12;
            this.checkPatID.Text = "Patient Id";
            this.checkPatID.UseVisualStyleBackColor = true;
            this.checkPatID.CheckedChanged += new System.EventHandler(this.checkPatID_CheckedChanged);
            // 
            // comboMLCLike
            // 
            this.comboMLCLike.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboMLCLike.FormattingEnabled = true;
            this.comboMLCLike.Items.AddRange(new object[] {
            "like",
            "not like"});
            this.comboMLCLike.Location = new System.Drawing.Point(123, 45);
            this.comboMLCLike.Name = "comboMLCLike";
            this.comboMLCLike.Size = new System.Drawing.Size(75, 21);
            this.comboMLCLike.TabIndex = 4;
            // 
            // comboGantLike
            // 
            this.comboGantLike.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboGantLike.FormattingEnabled = true;
            this.comboGantLike.Items.AddRange(new object[] {
            "like",
            "not like"});
            this.comboGantLike.Location = new System.Drawing.Point(123, 20);
            this.comboGantLike.Name = "comboGantLike";
            this.comboGantLike.Size = new System.Drawing.Size(75, 21);
            this.comboGantLike.TabIndex = 1;
            // 
            // numericFract
            // 
            this.numericFract.Location = new System.Drawing.Point(162, 96);
            this.numericFract.Name = "numericFract";
            this.numericFract.Size = new System.Drawing.Size(40, 20);
            this.numericFract.TabIndex = 11;
            this.numericFract.ValueChanged += new System.EventHandler(this.numericUpDown2_ValueChanged);
            // 
            // numericDose
            // 
            this.numericDose.Location = new System.Drawing.Point(162, 71);
            this.numericDose.Name = "numericDose";
            this.numericDose.Size = new System.Drawing.Size(40, 20);
            this.numericDose.TabIndex = 8;
            this.numericDose.ValueChanged += new System.EventHandler(this.numericUpDown1_ValueChanged);
            // 
            // comboFractEq
            // 
            this.comboFractEq.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboFractEq.FormattingEnabled = true;
            this.comboFractEq.Items.AddRange(new object[] {
            ">",
            "<",
            "="});
            this.comboFractEq.Location = new System.Drawing.Point(123, 95);
            this.comboFractEq.Name = "comboFractEq";
            this.comboFractEq.Size = new System.Drawing.Size(33, 21);
            this.comboFractEq.TabIndex = 10;
            // 
            // comboDoseEq
            // 
            this.comboDoseEq.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboDoseEq.FormattingEnabled = true;
            this.comboDoseEq.Items.AddRange(new object[] {
            ">",
            "<",
            "="});
            this.comboDoseEq.Location = new System.Drawing.Point(123, 70);
            this.comboDoseEq.Name = "comboDoseEq";
            this.comboDoseEq.Size = new System.Drawing.Size(33, 21);
            this.comboDoseEq.TabIndex = 7;
            // 
            // comboMLCCon
            // 
            this.comboMLCCon.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboMLCCon.FormattingEnabled = true;
            this.comboMLCCon.Items.AddRange(new object[] {
            "StdMLCPlan",
            "DynMLCPlan"});
            this.comboMLCCon.Location = new System.Drawing.Point(204, 45);
            this.comboMLCCon.Name = "comboMLCCon";
            this.comboMLCCon.Size = new System.Drawing.Size(80, 21);
            this.comboMLCCon.TabIndex = 5;
            // 
            // comboGantryCond
            // 
            this.comboGantryCond.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboGantryCond.FormattingEnabled = true;
            this.comboGantryCond.Items.AddRange(new object[] {
            "CW",
            "CCW",
            "NONE"});
            this.comboGantryCond.Location = new System.Drawing.Point(204, 20);
            this.comboGantryCond.Name = "comboGantryCond";
            this.comboGantryCond.Size = new System.Drawing.Size(80, 21);
            this.comboGantryCond.TabIndex = 2;
            // 
            // checkNFract
            // 
            this.checkNFract.AutoSize = true;
            this.checkNFract.Location = new System.Drawing.Point(17, 98);
            this.checkNFract.Name = "checkNFract";
            this.checkNFract.Size = new System.Drawing.Size(101, 17);
            this.checkNFract.TabIndex = 9;
            this.checkNFract.Text = "No. of Fractions";
            this.checkNFract.UseVisualStyleBackColor = true;
            this.checkNFract.CheckedChanged += new System.EventHandler(this.checkNFract_CheckedChanged);
            // 
            // checkPresDose
            // 
            this.checkPresDose.AutoSize = true;
            this.checkPresDose.Location = new System.Drawing.Point(17, 74);
            this.checkPresDose.Name = "checkPresDose";
            this.checkPresDose.Size = new System.Drawing.Size(104, 17);
            this.checkPresDose.TabIndex = 6;
            this.checkPresDose.Text = "Prescribed Dose";
            this.checkPresDose.UseVisualStyleBackColor = true;
            this.checkPresDose.CheckedChanged += new System.EventHandler(this.checkPresDose_CheckedChanged);
            // 
            // checkMLCPlan
            // 
            this.checkMLCPlan.AutoSize = true;
            this.checkMLCPlan.Location = new System.Drawing.Point(17, 50);
            this.checkMLCPlan.Name = "checkMLCPlan";
            this.checkMLCPlan.Size = new System.Drawing.Size(99, 17);
            this.checkMLCPlan.TabIndex = 3;
            this.checkMLCPlan.Text = "MLC Plan Type";
            this.checkMLCPlan.UseVisualStyleBackColor = true;
            this.checkMLCPlan.CheckedChanged += new System.EventHandler(this.checkMLCPlan_CheckedChanged);
            // 
            // checkGantryRot
            // 
            this.checkGantryRot.AutoSize = true;
            this.checkGantryRot.Location = new System.Drawing.Point(17, 25);
            this.checkGantryRot.Name = "checkGantryRot";
            this.checkGantryRot.Size = new System.Drawing.Size(100, 17);
            this.checkGantryRot.TabIndex = 0;
            this.checkGantryRot.Text = "Gantry Rotation";
            this.checkGantryRot.UseVisualStyleBackColor = true;
            this.checkGantryRot.CheckedChanged += new System.EventHandler(this.checkGantryRot_CheckedChanged);
            // 
            // linkLabelCancel
            // 
            this.linkLabelCancel.AutoSize = true;
            this.linkLabelCancel.Enabled = false;
            this.linkLabelCancel.Location = new System.Drawing.Point(132, 472);
            this.linkLabelCancel.Name = "linkLabelCancel";
            this.linkLabelCancel.Size = new System.Drawing.Size(40, 13);
            this.linkLabelCancel.TabIndex = 8;
            this.linkLabelCancel.TabStop = true;
            this.linkLabelCancel.Text = "Cancel";
            this.linkLabelCancel.Visible = false;
            this.linkLabelCancel.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabelBaseQ_LinkClicked);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(584, 562);
            this.Controls.Add(this.linkLabelCancel);
            this.Controls.Add(this.groupBoxAdvQ);
            this.Controls.Add(this.linkLabelAdvQ);
            this.Controls.Add(this.buttonExtract);
            this.Controls.Add(this.groupBoxQuery);
            this.Controls.Add(this.textBoxExcel);
            this.Controls.Add(this.groupBoxDatabase);
            this.Controls.Add(this.labelExcel);
            this.Controls.Add(this.buttonQuery);
            this.MaximumSize = new System.Drawing.Size(600, 600);
            this.MinimumSize = new System.Drawing.Size(333, 550);
            this.Name = "Form1";
            this.Text = "Database Query";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.groupBoxDatabase.ResumeLayout(false);
            this.groupBoxDatabase.PerformLayout();
            this.groupBoxQuery.ResumeLayout(false);
            this.groupBoxQuery.PerformLayout();
            this.groupBoxAdvQ.ResumeLayout(false);
            this.groupBoxAdvQ.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numericFract)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericDose)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.Button buttonQuery;
		private System.Windows.Forms.GroupBox groupBoxDatabase;
		private System.Windows.Forms.TextBox textBoxPassword;
		private System.Windows.Forms.TextBox textBoxUserID;
		private System.Windows.Forms.TextBox textBoxDatabase;
		private System.Windows.Forms.Label labelPassword;
		private System.Windows.Forms.Label labelUserID;
		private System.Windows.Forms.Label labelDatabase;
		private System.Windows.Forms.Label labelServer;
		private System.Windows.Forms.TextBox textBoxServer;
		private System.Windows.Forms.GroupBox groupBoxQuery;
		private System.Windows.Forms.TextBox textBoxExcel;
		private System.Windows.Forms.Label labelExcel;
		private System.Windows.Forms.Label labelStructure;
		private System.Windows.Forms.Label labelEnd;
		private System.Windows.Forms.Label labelStart;
		private System.Windows.Forms.ComboBox comboBoxSite;
		private System.Windows.Forms.DateTimePicker dateTimePickerStart;
		private System.Windows.Forms.DateTimePicker dateTimePickerEnd;
		private System.Windows.Forms.Button buttonExtract;
		private System.Windows.Forms.TextBox textBoxVolCode;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.LinkLabel linkLabelAdvQ;
		private System.Windows.Forms.GroupBox groupBoxAdvQ;
		private System.Windows.Forms.LinkLabel linkLabelCancel;
		private System.Windows.Forms.CheckBox checkNFract;
		private System.Windows.Forms.CheckBox checkPresDose;
		private System.Windows.Forms.CheckBox checkMLCPlan;
		private System.Windows.Forms.CheckBox checkGantryRot;
		private System.Windows.Forms.ComboBox comboGantryCond;
		private System.Windows.Forms.ComboBox comboMLCCon;
		private System.Windows.Forms.ComboBox comboFractEq;
		private System.Windows.Forms.ComboBox comboDoseEq;
		private System.Windows.Forms.NumericUpDown numericFract;
		private System.Windows.Forms.NumericUpDown numericDose;
		private System.Windows.Forms.ComboBox comboMLCLike;
		private System.Windows.Forms.ComboBox comboGantLike;
        private System.Windows.Forms.ComboBox comboPatIDLike;
        private System.Windows.Forms.CheckBox checkPatID;
        private System.Windows.Forms.TextBox textPatIDCond;
    }
}

