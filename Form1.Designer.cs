namespace OCDataImporter
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.comboBoxSE = new System.Windows.Forms.ComboBox();
            this.comboBoxCRF = new System.Windows.Forms.ComboBox();
            this.comboBoxIT = new System.Windows.Forms.ComboBox();
            this.comboBoxGR = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.textBoxInput = new System.Windows.Forms.TextBox();
            this.buttonBrowse = new System.Windows.Forms.Button();
            this.buttonStartConversion = new System.Windows.Forms.Button();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.textBoxOutput = new System.Windows.Forms.TextBox();
            this.buttonCancel = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.buttonExit = new System.Windows.Forms.Button();
            this.button_start = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.DGStudyDataCol = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.DGOCTargetItem = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.DGIsSubjectId = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.DGIsDate = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.DGSex = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.DGPID = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.DGDOB = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.DGSTD = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.CopyCurrentTarget = new System.Windows.Forms.DataGridViewLinkColumn();
            this.label7 = new System.Windows.Forms.Label();
            this.comboBoxDateFormat = new System.Windows.Forms.ComboBox();
            this.label8 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.comboBoxSex = new System.Windows.Forms.ComboBox();
            this.label10 = new System.Windows.Forms.Label();
            this.linkLabelBuildDG = new System.Windows.Forms.LinkLabel();
            this.label11 = new System.Windows.Forms.Label();
            this.textBoxMaxLines = new System.Windows.Forms.TextBox();
            this.label12 = new System.Windows.Forms.Label();
            this.label13 = new System.Windows.Forms.Label();
            this.textBoxSubjectSexM = new System.Windows.Forms.TextBox();
            this.textBoxSubjectSexF = new System.Windows.Forms.TextBox();
            this.label14 = new System.Windows.Forms.Label();
            this.textBoxLocation = new System.Windows.Forms.TextBox();
            this.checkBoxDup = new System.Windows.Forms.CheckBox();
            this.linkLabel1 = new System.Windows.Forms.LinkLabel();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.radioButtonNoEVT = new System.Windows.Forms.RadioButton();
            this.radioButtonUseTD = new System.Windows.Forms.RadioButton();
            this.labelWarningCounter = new System.Windows.Forms.Label();
            this.label15 = new System.Windows.Forms.Label();
            this.label16 = new System.Windows.Forms.Label();
            this.label17 = new System.Windows.Forms.Label();
            this.buttonConfPars = new System.Windows.Forms.Button();
            this.buttonBackToBegin = new System.Windows.Forms.Button();
            this.linkbuttonSHCols = new System.Windows.Forms.LinkLabel();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // comboBoxSE
            // 
            this.comboBoxSE.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxSE.FormattingEnabled = true;
            this.comboBoxSE.Location = new System.Drawing.Point(21, 563);
            this.comboBoxSE.Name = "comboBoxSE";
            this.comboBoxSE.Size = new System.Drawing.Size(209, 21);
            this.comboBoxSE.Sorted = true;
            this.comboBoxSE.TabIndex = 40;
            this.comboBoxSE.SelectedIndexChanged += new System.EventHandler(this.comboBoxSE_SelectedIndexChanged);
            // 
            // comboBoxCRF
            // 
            this.comboBoxCRF.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxCRF.DropDownWidth = 161;
            this.comboBoxCRF.FormattingEnabled = true;
            this.comboBoxCRF.Location = new System.Drawing.Point(236, 563);
            this.comboBoxCRF.Name = "comboBoxCRF";
            this.comboBoxCRF.Size = new System.Drawing.Size(195, 21);
            this.comboBoxCRF.Sorted = true;
            this.comboBoxCRF.TabIndex = 50;
            this.comboBoxCRF.SelectedIndexChanged += new System.EventHandler(this.comboBoxCRF_SelectedIndexChanged);
            // 
            // comboBoxIT
            // 
            this.comboBoxIT.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxIT.DropDownWidth = 161;
            this.comboBoxIT.FormattingEnabled = true;
            this.comboBoxIT.Location = new System.Drawing.Point(660, 563);
            this.comboBoxIT.Name = "comboBoxIT";
            this.comboBoxIT.Size = new System.Drawing.Size(562, 21);
            this.comboBoxIT.Sorted = true;
            this.comboBoxIT.TabIndex = 70;
            this.comboBoxIT.SelectedIndexChanged += new System.EventHandler(this.comboBoxIT_SelectedIndexChanged);
            // 
            // comboBoxGR
            // 
            this.comboBoxGR.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxGR.DropDownWidth = 161;
            this.comboBoxGR.FormattingEnabled = true;
            this.comboBoxGR.Location = new System.Drawing.Point(438, 563);
            this.comboBoxGR.Name = "comboBoxGR";
            this.comboBoxGR.Size = new System.Drawing.Size(216, 21);
            this.comboBoxGR.Sorted = true;
            this.comboBoxGR.TabIndex = 60;
            this.comboBoxGR.SelectedIndexChanged += new System.EventHandler(this.comboBoxGR_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Verdana", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(41, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(0, 18);
            this.label1.TabIndex = 0;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(104, 33);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(812, 13);
            this.label3.TabIndex = 2;
            this.label3.Text = "OCMetaData file (XML), Study Data file (TXT) and label-oid file (OID), separated " +
    "by a \';\' or use \'Browse\' button. The label-oid file is optional.";
            // 
            // textBoxInput
            // 
            this.textBoxInput.Location = new System.Drawing.Point(18, 52);
            this.textBoxInput.Name = "textBoxInput";
            this.textBoxInput.Size = new System.Drawing.Size(891, 21);
            this.textBoxInput.TabIndex = 3;
            // 
            // buttonBrowse
            // 
            this.buttonBrowse.Location = new System.Drawing.Point(916, 51);
            this.buttonBrowse.Name = "buttonBrowse";
            this.buttonBrowse.Size = new System.Drawing.Size(75, 23);
            this.buttonBrowse.TabIndex = 4;
            this.buttonBrowse.Text = "Browse";
            this.buttonBrowse.UseVisualStyleBackColor = true;
            this.buttonBrowse.Click += new System.EventHandler(this.buttonBrowse_Click_1);
            // 
            // buttonStartConversion
            // 
            this.buttonStartConversion.BackColor = System.Drawing.SystemColors.Control;
            this.buttonStartConversion.Enabled = false;
            this.buttonStartConversion.Location = new System.Drawing.Point(726, 592);
            this.buttonStartConversion.Name = "buttonStartConversion";
            this.buttonStartConversion.Size = new System.Drawing.Size(80, 23);
            this.buttonStartConversion.TabIndex = 7;
            this.buttonStartConversion.Text = "Start";
            this.buttonStartConversion.UseVisualStyleBackColor = false;
            this.buttonStartConversion.Click += new System.EventHandler(this.buttonStartConversion_Click_1);
            // 
            // progressBar1
            // 
            this.progressBar1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.progressBar1.BackColor = System.Drawing.SystemColors.Control;
            this.progressBar1.Location = new System.Drawing.Point(104, 621);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(1117, 23);
            this.progressBar1.TabIndex = 8;
            // 
            // textBoxOutput
            // 
            this.textBoxOutput.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxOutput.Location = new System.Drawing.Point(44, 653);
            this.textBoxOutput.Multiline = true;
            this.textBoxOutput.Name = "textBoxOutput";
            this.textBoxOutput.ReadOnly = true;
            this.textBoxOutput.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.textBoxOutput.Size = new System.Drawing.Size(1178, 85);
            this.textBoxOutput.TabIndex = 9;
            // 
            // buttonCancel
            // 
            this.buttonCancel.Enabled = false;
            this.buttonCancel.Location = new System.Drawing.Point(820, 591);
            this.buttonCancel.Name = "buttonCancel";
            this.buttonCancel.Size = new System.Drawing.Size(75, 23);
            this.buttonCancel.TabIndex = 10;
            this.buttonCancel.Text = "Cancel";
            this.buttonCancel.UseVisualStyleBackColor = true;
            this.buttonCancel.Click += new System.EventHandler(this.buttonCancel_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(41, 626);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(57, 13);
            this.label4.TabIndex = 11;
            this.label4.Text = "Progress";
            // 
            // buttonExit
            // 
            this.buttonExit.Location = new System.Drawing.Point(1143, 52);
            this.buttonExit.Name = "buttonExit";
            this.buttonExit.Size = new System.Drawing.Size(75, 23);
            this.buttonExit.TabIndex = 12;
            this.buttonExit.Text = "Exit";
            this.buttonExit.UseVisualStyleBackColor = true;
            this.buttonExit.Click += new System.EventHandler(this.buttonExit_Click);
            // 
            // button_start
            // 
            this.button_start.Location = new System.Drawing.Point(997, 51);
            this.button_start.Name = "button_start";
            this.button_start.Size = new System.Drawing.Size(135, 23);
            this.button_start.TabIndex = 61;
            this.button_start.Text = "Read Input Files";
            this.button_start.UseVisualStyleBackColor = true;
            this.button_start.Click += new System.EventHandler(this.button_start_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(114, 547);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(76, 13);
            this.label2.TabIndex = 62;
            this.label2.Text = "Study Event";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(233, 547);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(30, 13);
            this.label5.TabIndex = 63;
            this.label5.Text = "CRF";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(435, 547);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(42, 13);
            this.label6.TabIndex = 64;
            this.label6.Text = "Group";
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToOrderColumns = true;
            this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.DGStudyDataCol,
            this.DGOCTargetItem,
            this.DGIsSubjectId,
            this.DGIsDate,
            this.DGSex,
            this.DGPID,
            this.DGDOB,
            this.DGSTD,
            this.CopyCurrentTarget});
            this.dataGridView1.Location = new System.Drawing.Point(18, 196);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(1203, 338);
            this.dataGridView1.TabIndex = 13;
            this.dataGridView1.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellContentClick);
            // 
            // DGStudyDataCol
            // 
            this.DGStudyDataCol.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.DGStudyDataCol.HeaderText = "Study Data Column";
            this.DGStudyDataCol.Name = "DGStudyDataCol";
            this.DGStudyDataCol.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.DGStudyDataCol.ToolTipText = "Study data colomns based on the first line of data file";
            // 
            // DGOCTargetItem
            // 
            this.DGOCTargetItem.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.DGOCTargetItem.HeaderText = "OC Target Item";
            this.DGOCTargetItem.Name = "DGOCTargetItem";
            // 
            // DGIsSubjectId
            // 
            this.DGIsSubjectId.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.DisplayedCells;
            this.DGIsSubjectId.HeaderText = "Study Subject ID?";
            this.DGIsSubjectId.Name = "DGIsSubjectId";
            this.DGIsSubjectId.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.DGIsSubjectId.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.DGIsSubjectId.ToolTipText = "Is this field (part of) subject id?";
            this.DGIsSubjectId.Width = 106;
            // 
            // DGIsDate
            // 
            this.DGIsDate.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.ColumnHeader;
            this.DGIsDate.HeaderText = "Date?";
            this.DGIsDate.Name = "DGIsDate";
            this.DGIsDate.ToolTipText = "Is this field a date field?";
            this.DGIsDate.Visible = false;
            // 
            // DGSex
            // 
            this.DGSex.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.ColumnHeader;
            this.DGSex.HeaderText = "Subject Sex?";
            this.DGSex.Name = "DGSex";
            this.DGSex.ToolTipText = "Is this field the Subject\'s sex?";
            this.DGSex.Width = 79;
            // 
            // DGPID
            // 
            this.DGPID.HeaderText = "Subject Person ID?";
            this.DGPID.Name = "DGPID";
            this.DGPID.Width = 71;
            // 
            // DGDOB
            // 
            this.DGDOB.HeaderText = "Subject Date of Birth?";
            this.DGDOB.Name = "DGDOB";
            this.DGDOB.Width = 71;
            // 
            // DGSTD
            // 
            this.DGSTD.HeaderText = "Subject start date?";
            this.DGSTD.Name = "DGSTD";
            this.DGSTD.Width = 71;
            // 
            // CopyCurrentTarget
            // 
            this.CopyCurrentTarget.HeaderText = "CopyTarget";
            this.CopyCurrentTarget.Name = "CopyCurrentTarget";
            this.CopyCurrentTarget.Text = "CopyTarget";
            this.CopyCurrentTarget.UseColumnTextForLinkValue = true;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(657, 547);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(34, 13);
            this.label7.TabIndex = 71;
            this.label7.Text = "Item";
            // 
            // comboBoxDateFormat
            // 
            this.comboBoxDateFormat.FormattingEnabled = true;
            this.comboBoxDateFormat.Items.AddRange(new object[] {
            "--select--",
            "day-month-year",
            "month-day-year",
            "year-month-day"});
            this.comboBoxDateFormat.Location = new System.Drawing.Point(183, 101);
            this.comboBoxDateFormat.Name = "comboBoxDateFormat";
            this.comboBoxDateFormat.Size = new System.Drawing.Size(121, 21);
            this.comboBoxDateFormat.TabIndex = 72;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(17, 104);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(160, 13);
            this.label8.TabIndex = 73;
            this.label8.Text = "Date format in study items";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(18, 163);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(120, 13);
            this.label9.TabIndex = 75;
            this.label9.Text = "Replace string pairs";
            this.label9.Visible = false;
            // 
            // comboBoxSex
            // 
            this.comboBoxSex.FormattingEnabled = true;
            this.comboBoxSex.Items.AddRange(new object[] {
            "f",
            "m"});
            this.comboBoxSex.Location = new System.Drawing.Point(461, 101);
            this.comboBoxSex.Name = "comboBoxSex";
            this.comboBoxSex.Size = new System.Drawing.Size(60, 21);
            this.comboBoxSex.TabIndex = 76;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(316, 104);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(138, 13);
            this.label10.TabIndex = 77;
            this.label10.Text = "Default sex of subjects";
            // 
            // linkLabelBuildDG
            // 
            this.linkLabelBuildDG.AutoSize = true;
            this.linkLabelBuildDG.Location = new System.Drawing.Point(486, 597);
            this.linkLabelBuildDG.Name = "linkLabelBuildDG";
            this.linkLabelBuildDG.Size = new System.Drawing.Size(91, 13);
            this.linkLabelBuildDG.TabIndex = 78;
            this.linkLabelBuildDG.TabStop = true;
            this.linkLabelBuildDG.Text = "Match columns";
            this.linkLabelBuildDG.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabelBuildDG_LinkClicked);
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(16, 136);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(504, 13);
            this.label11.TabIndex = 79;
            this.label11.Text = "Split the ODM file where each contains  the following number of subjects (0 = no " +
    "split) ";
            // 
            // textBoxMaxLines
            // 
            this.textBoxMaxLines.Location = new System.Drawing.Point(526, 133);
            this.textBoxMaxLines.Name = "textBoxMaxLines";
            this.textBoxMaxLines.Size = new System.Drawing.Size(51, 21);
            this.textBoxMaxLines.TabIndex = 77;
            this.textBoxMaxLines.Text = "0";
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(527, 104);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(127, 13);
            this.label12.TabIndex = 80;
            this.label12.Text = "Gender Code for m: ";
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(723, 104);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(116, 13);
            this.label13.TabIndex = 81;
            this.label13.Text = "Gender Code for f:";
            // 
            // textBoxSubjectSexM
            // 
            this.textBoxSubjectSexM.Location = new System.Drawing.Point(660, 101);
            this.textBoxSubjectSexM.Name = "textBoxSubjectSexM";
            this.textBoxSubjectSexM.Size = new System.Drawing.Size(57, 21);
            this.textBoxSubjectSexM.TabIndex = 82;
            // 
            // textBoxSubjectSexF
            // 
            this.textBoxSubjectSexF.Location = new System.Drawing.Point(845, 101);
            this.textBoxSubjectSexF.Name = "textBoxSubjectSexF";
            this.textBoxSubjectSexF.Size = new System.Drawing.Size(50, 21);
            this.textBoxSubjectSexF.TabIndex = 83;
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Location = new System.Drawing.Point(595, 136);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(59, 13);
            this.label14.TabIndex = 84;
            this.label14.Text = "Location:";
            // 
            // textBoxLocation
            // 
            this.textBoxLocation.Location = new System.Drawing.Point(660, 133);
            this.textBoxLocation.Name = "textBoxLocation";
            this.textBoxLocation.Size = new System.Drawing.Size(151, 21);
            this.textBoxLocation.TabIndex = 85;
            // 
            // checkBoxDup
            // 
            this.checkBoxDup.AutoSize = true;
            this.checkBoxDup.Location = new System.Drawing.Point(909, 157);
            this.checkBoxDup.Name = "checkBoxDup";
            this.checkBoxDup.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.checkBoxDup.Size = new System.Drawing.Size(224, 17);
            this.checkBoxDup.TabIndex = 87;
            this.checkBoxDup.Text = "Check duplicate study subject ID\'s";
            this.checkBoxDup.TextAlign = System.Drawing.ContentAlignment.TopLeft;
            this.checkBoxDup.UseVisualStyleBackColor = true;
            // 
            // linkLabel1
            // 
            this.linkLabel1.AutoSize = true;
            this.linkLabel1.Location = new System.Drawing.Point(595, 596);
            this.linkLabel1.Name = "linkLabel1";
            this.linkLabel1.Size = new System.Drawing.Size(108, 13);
            this.linkLabel1.TabIndex = 91;
            this.linkLabel1.TabStop = true;
            this.linkLabel1.Text = "Unmatch columns";
            this.linkLabel1.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel1_LinkClicked);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.radioButtonNoEVT);
            this.groupBox1.Controls.Add(this.radioButtonUseTD);
            this.groupBox1.Location = new System.Drawing.Point(907, 94);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(304, 55);
            this.groupBox1.TabIndex = 92;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "If Subject start date is empty in data file:";
            // 
            // radioButtonNoEVT
            // 
            this.radioButtonNoEVT.AutoSize = true;
            this.radioButtonNoEVT.Location = new System.Drawing.Point(12, 34);
            this.radioButtonNoEVT.Name = "radioButtonNoEVT";
            this.radioButtonNoEVT.Size = new System.Drawing.Size(201, 17);
            this.radioButtonNoEVT.TabIndex = 1;
            this.radioButtonNoEVT.TabStop = true;
            this.radioButtonNoEVT.Text = "Do not generate Event records";
            this.radioButtonNoEVT.UseVisualStyleBackColor = true;
            // 
            // radioButtonUseTD
            // 
            this.radioButtonUseTD.AutoSize = true;
            this.radioButtonUseTD.Location = new System.Drawing.Point(12, 16);
            this.radioButtonUseTD.Name = "radioButtonUseTD";
            this.radioButtonUseTD.Size = new System.Drawing.Size(117, 17);
            this.radioButtonUseTD.TabIndex = 0;
            this.radioButtonUseTD.TabStop = true;
            this.radioButtonUseTD.Text = "Use todays date";
            this.radioButtonUseTD.UseVisualStyleBackColor = true;
            // 
            // labelWarningCounter
            // 
            this.labelWarningCounter.AutoSize = true;
            this.labelWarningCounter.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
            this.labelWarningCounter.Location = new System.Drawing.Point(51, 596);
            this.labelWarningCounter.Name = "labelWarningCounter";
            this.labelWarningCounter.Size = new System.Drawing.Size(0, 13);
            this.labelWarningCounter.TabIndex = 93;
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label15.Location = new System.Drawing.Point(18, 82);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(147, 13);
            this.label15.TabIndex = 94;
            this.label15.Text = "Program Parametres:";
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label16.Location = new System.Drawing.Point(18, 33);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(81, 13);
            this.label16.TabIndex = 95;
            this.label16.Text = "Input Files:";
            // 
            // label17
            // 
            this.label17.AutoSize = true;
            this.label17.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label17.Location = new System.Drawing.Point(17, 547);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(75, 13);
            this.label17.TabIndex = 96;
            this.label17.Text = "OC Target:";
            // 
            // buttonConfPars
            // 
            this.buttonConfPars.Location = new System.Drawing.Point(18, 157);
            this.buttonConfPars.Name = "buttonConfPars";
            this.buttonConfPars.Size = new System.Drawing.Size(224, 23);
            this.buttonConfPars.TabIndex = 97;
            this.buttonConfPars.Text = "Confirm Program Parametres";
            this.buttonConfPars.UseVisualStyleBackColor = true;
            this.buttonConfPars.Click += new System.EventHandler(this.buttonConfPars_Click);
            // 
            // buttonBackToBegin
            // 
            this.buttonBackToBegin.Location = new System.Drawing.Point(919, 590);
            this.buttonBackToBegin.Name = "buttonBackToBegin";
            this.buttonBackToBegin.Size = new System.Drawing.Size(142, 23);
            this.buttonBackToBegin.TabIndex = 98;
            this.buttonBackToBegin.Text = "Back To Beginning";
            this.buttonBackToBegin.UseVisualStyleBackColor = true;
            this.buttonBackToBegin.Click += new System.EventHandler(this.buttonBackToBegin_Click);
            // 
            // linkbuttonSHCols
            // 
            this.linkbuttonSHCols.AutoSize = true;
            this.linkbuttonSHCols.Location = new System.Drawing.Point(274, 597);
            this.linkbuttonSHCols.Name = "linkbuttonSHCols";
            this.linkbuttonSHCols.Size = new System.Drawing.Size(180, 13);
            this.linkbuttonSHCols.TabIndex = 99;
            this.linkbuttonSHCols.TabStop = true;
            this.linkbuttonSHCols.Text = "Hide Subject Related Columns";
            this.linkbuttonSHCols.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkbuttonSHCols_LinkClicked);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1247, 750);
            this.Controls.Add(this.linkbuttonSHCols);
            this.Controls.Add(this.buttonBackToBegin);
            this.Controls.Add(this.buttonConfPars);
            this.Controls.Add(this.label17);
            this.Controls.Add(this.label16);
            this.Controls.Add(this.label15);
            this.Controls.Add(this.labelWarningCounter);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.linkLabel1);
            this.Controls.Add(this.checkBoxDup);
            this.Controls.Add(this.textBoxLocation);
            this.Controls.Add(this.label14);
            this.Controls.Add(this.textBoxSubjectSexF);
            this.Controls.Add(this.textBoxSubjectSexM);
            this.Controls.Add(this.label13);
            this.Controls.Add(this.label12);
            this.Controls.Add(this.textBoxMaxLines);
            this.Controls.Add(this.label11);
            this.Controls.Add(this.linkLabelBuildDG);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.comboBoxSex);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.comboBoxDateFormat);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.comboBoxIT);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.button_start);
            this.Controls.Add(this.buttonExit);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.buttonCancel);
            this.Controls.Add(this.textBoxOutput);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.buttonStartConversion);
            this.Controls.Add(this.buttonBrowse);
            this.Controls.Add(this.textBoxInput);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.comboBoxGR);
            this.Controls.Add(this.comboBoxCRF);
            this.Controls.Add(this.comboBoxSE);
            this.Controls.Add(this.dataGridView1);
            this.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.Text = "OCDataImporter - Provided by VU Medical Center, dept. of Pathology, Amsterdam";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox textBoxInput;
        private System.Windows.Forms.Button buttonBrowse;
        private System.Windows.Forms.Button buttonStartConversion;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.TextBox textBoxOutput;
        private System.Windows.Forms.ComboBox comboBoxIT; 
        private System.Windows.Forms.Button buttonCancel;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button buttonExit;
        private System.Windows.Forms.ComboBox comboBoxSE;
        private System.Windows.Forms.ComboBox comboBoxCRF;
        private System.Windows.Forms.ComboBox comboBoxGR;
        private System.Windows.Forms.Button button_start;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.ComboBox comboBoxDateFormat;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.ComboBox comboBoxSex;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.LinkLabel linkLabelBuildDG;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.TextBox textBoxMaxLines;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.TextBox textBoxSubjectSexM;
        private System.Windows.Forms.TextBox textBoxSubjectSexF;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.TextBox textBoxLocation;
        private System.Windows.Forms.CheckBox checkBoxDup;
        private System.Windows.Forms.LinkLabel linkLabel1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.RadioButton radioButtonNoEVT;
        private System.Windows.Forms.RadioButton radioButtonUseTD;
        private System.Windows.Forms.Label labelWarningCounter;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.Label label16;
        private System.Windows.Forms.Label label17;
        private System.Windows.Forms.Button buttonConfPars;
        private System.Windows.Forms.Button buttonBackToBegin;
        private System.Windows.Forms.DataGridViewTextBoxColumn DGStudyDataCol;
        private System.Windows.Forms.DataGridViewTextBoxColumn DGOCTargetItem;
        private System.Windows.Forms.DataGridViewCheckBoxColumn DGIsSubjectId;
        private System.Windows.Forms.DataGridViewCheckBoxColumn DGIsDate;
        private System.Windows.Forms.DataGridViewCheckBoxColumn DGSex;
        private System.Windows.Forms.DataGridViewCheckBoxColumn DGPID;
        private System.Windows.Forms.DataGridViewCheckBoxColumn DGDOB;
        private System.Windows.Forms.DataGridViewCheckBoxColumn DGSTD;
        private System.Windows.Forms.DataGridViewLinkColumn CopyCurrentTarget;
        private System.Windows.Forms.LinkLabel linkbuttonSHCols;
        private System.Windows.Forms.ToolTip toolTip1;
    }
}

