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
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.comboBoxCaller = new System.Windows.Forms.ComboBox();
            this.textBoxCallerPicture = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // buttonOk
            // 
            this.buttonOk.Location = new System.Drawing.Point(308, 313);
            this.buttonOk.Name = "buttonOk";
            this.buttonOk.Size = new System.Drawing.Size(143, 54);
            this.buttonOk.TabIndex = 0;
            this.buttonOk.Text = "OK";
            this.buttonOk.UseVisualStyleBackColor = true;
            this.buttonOk.Click += new System.EventHandler(this.buttonOk_Click);
            // 
            // textBoxDanceName
            // 
            this.textBoxDanceName.Location = new System.Drawing.Point(31, 33);
            this.textBoxDanceName.Name = "textBoxDanceName";
            this.textBoxDanceName.Size = new System.Drawing.Size(160, 20);
            this.textBoxDanceName.TabIndex = 1;
            this.textBoxDanceName.Text = "Jesperdansen";
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
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.radioButtonFestival);
            this.groupBox1.Controls.Add(this.radioButtonWeekEnd);
            this.groupBox1.Location = new System.Drawing.Point(261, 35);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(110, 79);
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
            this.groupBox2.Location = new System.Drawing.Point(21, 167);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(226, 93);
            this.groupBox2.TabIndex = 11;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Datum";
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
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(16, 55);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(25, 13);
            this.label4.TabIndex = 12;
            this.label4.Text = "Slut";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(34, 124);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(62, 13);
            this.label5.TabIndex = 13;
            this.label5.Text = "Caller namn";
            // 
            // comboBoxCaller
            // 
            this.comboBoxCaller.FormattingEnabled = true;
            this.comboBoxCaller.Location = new System.Drawing.Point(31, 140);
            this.comboBoxCaller.Name = "comboBoxCaller";
            this.comboBoxCaller.Size = new System.Drawing.Size(141, 21);
            this.comboBoxCaller.TabIndex = 14;
            this.comboBoxCaller.SelectedIndexChanged += new System.EventHandler(this.comboBoxCaller_SelectedIndexChanged);
            // 
            // textBoxCallerPicture
            // 
            this.textBoxCallerPicture.Location = new System.Drawing.Point(200, 141);
            this.textBoxCallerPicture.Name = "textBoxCallerPicture";
            this.textBoxCallerPicture.Size = new System.Drawing.Size(217, 20);
            this.textBoxCallerPicture.TabIndex = 15;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(197, 125);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(52, 13);
            this.label6.TabIndex = 16;
            this.label6.Text = "Caller bild";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(510, 391);
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
    }
}

