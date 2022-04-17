using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CreateWordFiles
{
    public partial class Form1 : Form
    {
        Dictionary<string, string> Caller_dictionary = new Dictionary<string, string>();
        public Form1()
        {
            InitializeComponent();
            this.myInit();
        }
        private void myInit()
        {
            this.folderBrowserDialog1.Description =
            "Välj katalog för utdata.";

            var lines = System.IO.File.ReadAllLines(@"Resources\Callers.txt");
            foreach (var line in lines)
            {
                String[] atoms= line.Split(new char[] { ';' });
                this.comboBoxCaller.Items.Add(atoms[0]);
                Caller_dictionary.Add(atoms[0], atoms[1]);  
            }
        }
        private void buttonOk_Click(object sender, EventArgs e)
        {
            //String file = @"d:\Development\Visual Studio\Projects\CreateWordFiles\CreateWordFiles\Documents\Test.docx";
           
            Creator.CreateWordprocessingDocument(this.textBoxOutputFolder.Text, this.textBoxDanceName.Text, this.dateTimePickerStart.Value, this.dateTimePickerEnd.Value);
            //System.Diagnostics.Process.Start(file);
            this.Close();
        }

        private void comboBoxCaller_SelectedIndexChanged(object sender, EventArgs e)
        {
            String callerName= comboBoxCaller.SelectedItem.ToString();
            String callerPicture = Caller_dictionary[callerName];
            this.textBoxCallerPicture.Text = callerPicture;
        }

        private void dateTimePickerStart_ValueChanged(object sender, EventArgs e)
        {
            System.TimeSpan ts = new TimeSpan(24, 0, 0);
            if (this.radioButtonFestival.Checked)
            {
                ts = new TimeSpan(24*3, 0, 0);

            }
            this.dateTimePickerEnd.Value = this.dateTimePickerStart.Value + ts;
        }

        private void buttonSelectOutputFolder_Click(object sender, EventArgs e)
        {

            DialogResult result = folderBrowserDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                this.textBoxOutputFolder.Text = folderBrowserDialog1.SelectedPath;
            }
        }
    }
}
