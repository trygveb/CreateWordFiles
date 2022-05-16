namespace CreateWordFiles
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
            this.buttonOk = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.dateTimePickerStart = new System.Windows.Forms.DateTimePicker();
            this.dateTimePickerEnd = new System.Windows.Forms.DateTimePicker();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.comboBoxCaller = new System.Windows.Forms.ComboBox();
            this.textBoxCallerPicture = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.textBoxOutputFolder = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.buttonSelectOutputFolder = new System.Windows.Forms.Button();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.buttonCancel = new System.Windows.Forms.Button();
            this.groupBoxLanguage = new System.Windows.Forms.GroupBox();
            this.radioButtonEnglish = new System.Windows.Forms.RadioButton();
            this.radioButtonSwedish = new System.Windows.Forms.RadioButton();
            this.comboBoxDanceSchema = new System.Windows.Forms.ComboBox();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.comboBoxDanceLocation = new System.Windows.Forms.ComboBox();
            this.groupBoxCoffee = new System.Windows.Forms.GroupBox();
            this.radioButtonCoffeeNo = new System.Windows.Forms.RadioButton();
            this.radioButtonCoffeeYes = new System.Windows.Forms.RadioButton();
            this.comboBoxDanceName = new System.Windows.Forms.ComboBox();
            this.groupBox2.SuspendLayout();
            this.groupBoxLanguage.SuspendLayout();
            this.groupBoxCoffee.SuspendLayout();
            this.SuspendLayout();
            // 
            // buttonOk
            // 
            this.buttonOk.Enabled = false;
            this.buttonOk.Location = new System.Drawing.Point(308, 333);
            this.buttonOk.Name = "buttonOk";
            this.buttonOk.Size = new System.Drawing.Size(143, 34);
            this.buttonOk.TabIndex = 0;
            this.buttonOk.Text = "Skapa flyer";
            this.buttonOk.UseVisualStyleBackColor = true;
            this.buttonOk.Click += new System.EventHandler(this.buttonOk_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(18, 16);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(78, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Dansens namn";
            // 
            // dateTimePickerStart
            // 
            this.dateTimePickerStart.Location = new System.Drawing.Point(59, 18);
            this.dateTimePickerStart.Name = "dateTimePickerStart";
            this.dateTimePickerStart.Size = new System.Drawing.Size(138, 20);
            this.dateTimePickerStart.TabIndex = 8;
            this.dateTimePickerStart.Value = new System.DateTime(2022, 10, 7, 0, 0, 0, 0);
            this.dateTimePickerStart.ValueChanged += new System.EventHandler(this.dateTimePickerStart_ValueChanged);
            // 
            // dateTimePickerEnd
            // 
            this.dateTimePickerEnd.Location = new System.Drawing.Point(59, 51);
            this.dateTimePickerEnd.Name = "dateTimePickerEnd";
            this.dateTimePickerEnd.Size = new System.Drawing.Size(138, 20);
            this.dateTimePickerEnd.TabIndex = 10;
            this.dateTimePickerEnd.Value = new System.DateTime(2022, 10, 10, 0, 0, 0, 0);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.label4);
            this.groupBox2.Controls.Add(this.label3);
            this.groupBox2.Controls.Add(this.dateTimePickerEnd);
            this.groupBox2.Controls.Add(this.dateTimePickerStart);
            this.groupBox2.Location = new System.Drawing.Point(21, 77);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(226, 93);
            this.groupBox2.TabIndex = 11;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Datum";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(16, 55);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(25, 13);
            this.label4.TabIndex = 12;
            this.label4.Text = "Slut";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(16, 22);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(29, 13);
            this.label3.TabIndex = 11;
            this.label3.Text = "Start";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(24, 215);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(62, 13);
            this.label5.TabIndex = 13;
            this.label5.Text = "Caller namn";
            this.label5.Click += new System.EventHandler(this.label5_Click);
            // 
            // comboBoxCaller
            // 
            this.comboBoxCaller.FormattingEnabled = true;
            this.comboBoxCaller.Location = new System.Drawing.Point(21, 231);
            this.comboBoxCaller.Name = "comboBoxCaller";
            this.comboBoxCaller.Size = new System.Drawing.Size(141, 21);
            this.comboBoxCaller.TabIndex = 14;
            this.comboBoxCaller.SelectedIndexChanged += new System.EventHandler(this.comboBoxCaller_SelectedIndexChanged);
            // 
            // textBoxCallerPicture
            // 
            this.textBoxCallerPicture.Location = new System.Drawing.Point(190, 232);
            this.textBoxCallerPicture.Name = "textBoxCallerPicture";
            this.textBoxCallerPicture.Size = new System.Drawing.Size(298, 20);
            this.textBoxCallerPicture.TabIndex = 15;
            this.textBoxCallerPicture.TextChanged += new System.EventHandler(this.textBoxCallerPicture_TextChanged);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(187, 216);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(52, 13);
            this.label6.TabIndex = 16;
            this.label6.Text = "Caller bild";
            this.label6.Click += new System.EventHandler(this.label6_Click);
            // 
            // textBoxOutputFolder
            // 
            this.textBoxOutputFolder.Location = new System.Drawing.Point(31, 278);
            this.textBoxOutputFolder.Name = "textBoxOutputFolder";
            this.textBoxOutputFolder.Size = new System.Drawing.Size(298, 20);
            this.textBoxOutputFolder.TabIndex = 17;
            this.textBoxOutputFolder.Text = "D:\\Mina dokument\\Sqd\\Motiv8s\\Flyers";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(28, 262);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(77, 13);
            this.label2.TabIndex = 18;
            this.label2.Text = "Utdata katalog";
            // 
            // buttonSelectOutputFolder
            // 
            this.buttonSelectOutputFolder.Location = new System.Drawing.Point(335, 274);
            this.buttonSelectOutputFolder.Name = "buttonSelectOutputFolder";
            this.buttonSelectOutputFolder.Size = new System.Drawing.Size(107, 27);
            this.buttonSelectOutputFolder.TabIndex = 19;
            this.buttonSelectOutputFolder.Text = "Välj annan katalog";
            this.buttonSelectOutputFolder.UseVisualStyleBackColor = true;
            this.buttonSelectOutputFolder.Click += new System.EventHandler(this.buttonSelectOutputFolder_Click);
            // 
            // buttonCancel
            // 
            this.buttonCancel.Location = new System.Drawing.Point(27, 333);
            this.buttonCancel.Name = "buttonCancel";
            this.buttonCancel.Size = new System.Drawing.Size(109, 27);
            this.buttonCancel.TabIndex = 20;
            this.buttonCancel.Text = "Avbryt";
            this.buttonCancel.UseVisualStyleBackColor = true;
            this.buttonCancel.Click += new System.EventHandler(this.buttonCancel_Click);
            // 
            // groupBoxLanguage
            // 
            this.groupBoxLanguage.Controls.Add(this.radioButtonEnglish);
            this.groupBoxLanguage.Controls.Add(this.radioButtonSwedish);
            this.groupBoxLanguage.Location = new System.Drawing.Point(278, 4);
            this.groupBoxLanguage.Name = "groupBoxLanguage";
            this.groupBoxLanguage.Size = new System.Drawing.Size(183, 49);
            this.groupBoxLanguage.TabIndex = 21;
            this.groupBoxLanguage.TabStop = false;
            this.groupBoxLanguage.Text = "Språk";
            // 
            // radioButtonEnglish
            // 
            this.radioButtonEnglish.AutoSize = true;
            this.radioButtonEnglish.Location = new System.Drawing.Point(84, 19);
            this.radioButtonEnglish.Name = "radioButtonEnglish";
            this.radioButtonEnglish.Size = new System.Drawing.Size(69, 17);
            this.radioButtonEnglish.TabIndex = 1;
            this.radioButtonEnglish.Tag = "en";
            this.radioButtonEnglish.Text = "Engelska";
            this.radioButtonEnglish.UseVisualStyleBackColor = true;
            // 
            // radioButtonSwedish
            // 
            this.radioButtonSwedish.AutoSize = true;
            this.radioButtonSwedish.Checked = true;
            this.radioButtonSwedish.Location = new System.Drawing.Point(11, 19);
            this.radioButtonSwedish.Name = "radioButtonSwedish";
            this.radioButtonSwedish.Size = new System.Drawing.Size(67, 17);
            this.radioButtonSwedish.TabIndex = 0;
            this.radioButtonSwedish.TabStop = true;
            this.radioButtonSwedish.Tag = "se";
            this.radioButtonSwedish.Text = "Svenska";
            this.radioButtonSwedish.UseVisualStyleBackColor = true;
            // 
            // comboBoxDanceSchema
            // 
            this.comboBoxDanceSchema.FormattingEnabled = true;
            this.comboBoxDanceSchema.Location = new System.Drawing.Point(274, 77);
            this.comboBoxDanceSchema.Name = "comboBoxDanceSchema";
            this.comboBoxDanceSchema.Size = new System.Drawing.Size(214, 21);
            this.comboBoxDanceSchema.TabIndex = 22;
            this.comboBoxDanceSchema.SelectedIndexChanged += new System.EventHandler(this.comboBoxDanceSchema_SelectedIndexChanged);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(271, 61);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(69, 13);
            this.label7.TabIndex = 23;
            this.label7.Text = "Dansschema";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(271, 115);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(54, 13);
            this.label8.TabIndex = 25;
            this.label8.Text = "Danslokal";
            // 
            // comboBoxDanceLocation
            // 
            this.comboBoxDanceLocation.FormattingEnabled = true;
            this.comboBoxDanceLocation.Location = new System.Drawing.Point(274, 131);
            this.comboBoxDanceLocation.Name = "comboBoxDanceLocation";
            this.comboBoxDanceLocation.Size = new System.Drawing.Size(214, 21);
            this.comboBoxDanceLocation.TabIndex = 24;
            // 
            // groupBoxCoffee
            // 
            this.groupBoxCoffee.Controls.Add(this.radioButtonCoffeeNo);
            this.groupBoxCoffee.Controls.Add(this.radioButtonCoffeeYes);
            this.groupBoxCoffee.Location = new System.Drawing.Point(274, 168);
            this.groupBoxCoffee.Name = "groupBoxCoffee";
            this.groupBoxCoffee.Size = new System.Drawing.Size(183, 49);
            this.groupBoxCoffee.TabIndex = 26;
            this.groupBoxCoffee.TabStop = false;
            this.groupBoxCoffee.Text = "Fikaservering";
            // 
            // radioButtonCoffeeNo
            // 
            this.radioButtonCoffeeNo.AutoSize = true;
            this.radioButtonCoffeeNo.Checked = true;
            this.radioButtonCoffeeNo.Location = new System.Drawing.Point(84, 19);
            this.radioButtonCoffeeNo.Name = "radioButtonCoffeeNo";
            this.radioButtonCoffeeNo.Size = new System.Drawing.Size(41, 17);
            this.radioButtonCoffeeNo.TabIndex = 1;
            this.radioButtonCoffeeNo.TabStop = true;
            this.radioButtonCoffeeNo.Tag = "en";
            this.radioButtonCoffeeNo.Text = "Nej";
            this.radioButtonCoffeeNo.UseVisualStyleBackColor = true;
            // 
            // radioButtonCoffeeYes
            // 
            this.radioButtonCoffeeYes.AutoSize = true;
            this.radioButtonCoffeeYes.Location = new System.Drawing.Point(11, 19);
            this.radioButtonCoffeeYes.Name = "radioButtonCoffeeYes";
            this.radioButtonCoffeeYes.Size = new System.Drawing.Size(36, 17);
            this.radioButtonCoffeeYes.TabIndex = 0;
            this.radioButtonCoffeeYes.Tag = "se";
            this.radioButtonCoffeeYes.Text = "Ja";
            this.radioButtonCoffeeYes.UseVisualStyleBackColor = true;
            // 
            // comboBoxDanceName
            // 
            this.comboBoxDanceName.FormattingEnabled = true;
            this.comboBoxDanceName.Items.AddRange(new object[] {
            "Jesperdansen",
            "Thomasdansen",
            "Majdansen",
            "Oktoberfestivalen",
            "Februarifestivalen"});
            this.comboBoxDanceName.Location = new System.Drawing.Point(21, 32);
            this.comboBoxDanceName.Name = "comboBoxDanceName";
            this.comboBoxDanceName.Size = new System.Drawing.Size(214, 21);
            this.comboBoxDanceName.TabIndex = 27;
            this.comboBoxDanceName.SelectedIndexChanged += new System.EventHandler(this.comboBoxDanceName_SelectedIndexChanged);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(510, 391);
            this.Controls.Add(this.comboBoxDanceName);
            this.Controls.Add(this.groupBoxCoffee);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.comboBoxDanceLocation);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.comboBoxDanceSchema);
            this.Controls.Add(this.groupBoxLanguage);
            this.Controls.Add(this.buttonCancel);
            this.Controls.Add(this.buttonSelectOutputFolder);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.textBoxOutputFolder);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.textBoxCallerPicture);
            this.Controls.Add(this.comboBoxCaller);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.buttonOk);
            this.Name = "Form1";
            this.Text = "Form1";
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBoxLanguage.ResumeLayout(false);
            this.groupBoxLanguage.PerformLayout();
            this.groupBoxCoffee.ResumeLayout(false);
            this.groupBoxCoffee.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button buttonOk;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DateTimePicker dateTimePickerStart;
        private System.Windows.Forms.DateTimePicker dateTimePickerEnd;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.ComboBox comboBoxCaller;
        private System.Windows.Forms.TextBox textBoxCallerPicture;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox textBoxOutputFolder;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button buttonSelectOutputFolder;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.Button buttonCancel;
        private System.Windows.Forms.GroupBox groupBoxLanguage;
        private System.Windows.Forms.RadioButton radioButtonEnglish;
        private System.Windows.Forms.RadioButton radioButtonSwedish;
        private System.Windows.Forms.ComboBox comboBoxDanceSchema;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.ComboBox comboBoxDanceLocation;
        private System.Windows.Forms.GroupBox groupBoxCoffee;
        private System.Windows.Forms.RadioButton radioButtonCoffeeNo;
        private System.Windows.Forms.RadioButton radioButtonCoffeeYes;
        private System.Windows.Forms.ComboBox comboBoxDanceName;
    }
}

