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
            this.textBoxDanceName = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.radioButtonWeekEnd = new System.Windows.Forms.RadioButton();
            this.radioButtonFestival = new System.Windows.Forms.RadioButton();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
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
            this.radioButtonSwedish = new System.Windows.Forms.RadioButton();
            this.radioButtonEnglish = new System.Windows.Forms.RadioButton();
            this.radioButtonMeeting = new System.Windows.Forms.RadioButton();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBoxLanguage.SuspendLayout();
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
            // textBoxDanceName
            // 
            this.textBoxDanceName.Location = new System.Drawing.Point(31, 33);
            this.textBoxDanceName.Name = "textBoxDanceName";
            this.textBoxDanceName.Size = new System.Drawing.Size(160, 20);
            this.textBoxDanceName.TabIndex = 1;
            this.textBoxDanceName.Text = "Test_Dans";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(28, 17);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(78, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Dansens namn";
            // 
            // radioButtonWeekEnd
            // 
            this.radioButtonWeekEnd.AutoSize = true;
            this.radioButtonWeekEnd.Checked = true;
            this.radioButtonWeekEnd.Location = new System.Drawing.Point(6, 25);
            this.radioButtonWeekEnd.Name = "radioButtonWeekEnd";
            this.radioButtonWeekEnd.Size = new System.Drawing.Size(72, 17);
            this.radioButtonWeekEnd.TabIndex = 5;
            this.radioButtonWeekEnd.TabStop = true;
            this.radioButtonWeekEnd.Text = "Weekend";
            this.radioButtonWeekEnd.UseVisualStyleBackColor = true;
            // 
            // radioButtonFestival
            // 
            this.radioButtonFestival.AutoSize = true;
            this.radioButtonFestival.Location = new System.Drawing.Point(6, 48);
            this.radioButtonFestival.Name = "radioButtonFestival";
            this.radioButtonFestival.Size = new System.Drawing.Size(61, 17);
            this.radioButtonFestival.TabIndex = 6;
            this.radioButtonFestival.Text = "Festival";
            this.radioButtonFestival.UseVisualStyleBackColor = true;
            this.radioButtonFestival.CheckedChanged += new System.EventHandler(this.radioButtonFestival_CheckedChanged);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.radioButtonMeeting);
            this.groupBox1.Controls.Add(this.radioButtonFestival);
            this.groupBox1.Controls.Add(this.radioButtonWeekEnd);
            this.groupBox1.Location = new System.Drawing.Point(268, 17);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(203, 79);
            this.groupBox1.TabIndex = 7;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Danstyp";
            // 
            // dateTimePickerStart
            // 
            this.dateTimePickerStart.Location = new System.Drawing.Point(59, 18);
            this.dateTimePickerStart.Name = "dateTimePickerStart";
            this.dateTimePickerStart.Size = new System.Drawing.Size(138, 20);
            this.dateTimePickerStart.TabIndex = 8;
            this.dateTimePickerStart.ValueChanged += new System.EventHandler(this.dateTimePickerStart_ValueChanged);
            // 
            // dateTimePickerEnd
            // 
            this.dateTimePickerEnd.Location = new System.Drawing.Point(59, 51);
            this.dateTimePickerEnd.Name = "dateTimePickerEnd";
            this.dateTimePickerEnd.Size = new System.Drawing.Size(138, 20);
            this.dateTimePickerEnd.TabIndex = 10;
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
            this.label5.Location = new System.Drawing.Point(24, 196);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(62, 13);
            this.label5.TabIndex = 13;
            this.label5.Text = "Caller namn";
            // 
            // comboBoxCaller
            // 
            this.comboBoxCaller.FormattingEnabled = true;
            this.comboBoxCaller.Location = new System.Drawing.Point(21, 212);
            this.comboBoxCaller.Name = "comboBoxCaller";
            this.comboBoxCaller.Size = new System.Drawing.Size(141, 21);
            this.comboBoxCaller.TabIndex = 14;
            this.comboBoxCaller.SelectedIndexChanged += new System.EventHandler(this.comboBoxCaller_SelectedIndexChanged);
            // 
            // textBoxCallerPicture
            // 
            this.textBoxCallerPicture.Location = new System.Drawing.Point(190, 213);
            this.textBoxCallerPicture.Name = "textBoxCallerPicture";
            this.textBoxCallerPicture.Size = new System.Drawing.Size(298, 20);
            this.textBoxCallerPicture.TabIndex = 15;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(187, 197);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(52, 13);
            this.label6.TabIndex = 16;
            this.label6.Text = "Caller bild";
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
            this.groupBoxLanguage.Location = new System.Drawing.Point(268, 99);
            this.groupBoxLanguage.Name = "groupBoxLanguage";
            this.groupBoxLanguage.Size = new System.Drawing.Size(110, 71);
            this.groupBoxLanguage.TabIndex = 21;
            this.groupBoxLanguage.TabStop = false;
            this.groupBoxLanguage.Text = "Språk";
            // 
            // radioButtonSwedish
            // 
            this.radioButtonSwedish.AutoSize = true;
            this.radioButtonSwedish.Checked = true;
            this.radioButtonSwedish.Location = new System.Drawing.Point(11, 19);
            this.radioButtonSwedish.Name = "radioButtonSwedish";
            this.radioButtonSwedish.Size = new System.Drawing.Size(67, 17);
            this.radioButtonSwedish.TabIndex = 0;
            this.radioButtonSwedish.Tag = "se";
            this.radioButtonSwedish.Text = "Svenska";
            this.radioButtonSwedish.UseVisualStyleBackColor = true;
            // 
            // radioButtonEnglish
            // 
            this.radioButtonEnglish.AutoSize = true;
            this.radioButtonEnglish.Location = new System.Drawing.Point(11, 44);
            this.radioButtonEnglish.Name = "radioButtonEnglish";
            this.radioButtonEnglish.Size = new System.Drawing.Size(69, 17);
            this.radioButtonEnglish.TabIndex = 1;
            this.radioButtonEnglish.Tag = "en";
            this.radioButtonEnglish.Text = "Engelska";
            this.radioButtonEnglish.UseVisualStyleBackColor = true;
            // 
            // radioButtonMeeting
            // 
            this.radioButtonMeeting.AutoSize = true;
            this.radioButtonMeeting.Location = new System.Drawing.Point(102, 25);
            this.radioButtonMeeting.Name = "radioButtonMeeting";
            this.radioButtonMeeting.Size = new System.Drawing.Size(63, 17);
            this.radioButtonMeeting.TabIndex = 7;
            this.radioButtonMeeting.Text = "Årsmöte";
            this.radioButtonMeeting.UseVisualStyleBackColor = true;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(510, 391);
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
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.textBoxDanceName);
            this.Controls.Add(this.buttonOk);
            this.Name = "Form1";
            this.Text = "Form1";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBoxLanguage.ResumeLayout(false);
            this.groupBoxLanguage.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button buttonOk;
        private System.Windows.Forms.TextBox textBoxDanceName;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.RadioButton radioButtonWeekEnd;
        private System.Windows.Forms.RadioButton radioButtonFestival;
        private System.Windows.Forms.GroupBox groupBox1;
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
        private System.Windows.Forms.RadioButton radioButtonMeeting;
        private System.Windows.Forms.GroupBox groupBoxLanguage;
        private System.Windows.Forms.RadioButton radioButtonEnglish;
        private System.Windows.Forms.RadioButton radioButtonSwedish;
    }
}

