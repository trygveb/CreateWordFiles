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
        String callerName= "";
        String callerPictureFile = "";
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
            this.dateTimePickerEnd.Value = this.dateTimePickerStart.Value + new TimeSpan(24, 0, 0);
        }
        private void setEndDate()
        {
            System.TimeSpan ts = new TimeSpan(24, 0, 0);
            if (this.radioButtonFestival.Checked)
            {
                ts = new TimeSpan(24 * 3, 0, 0);

            }
            this.dateTimePickerEnd.Value = this.dateTimePickerStart.Value + ts;
        }
        private void buttonOk_Click(object sender, EventArgs e)
        {
            var radioButtonLanguage = groupBoxLanguage.Controls.OfType<RadioButton>()
                   .Where(r => r.Checked).FirstOrDefault();
            String lang = (String)radioButtonLanguage.Tag;
            Dictionary<String, String> texts = this.getTexts(lang);
            Creator.CreateWordprocessingDocument(texts, lang, this.dateTimePickerStart.Value, this.dateTimePickerEnd.Value);
            //System.Diagnostics.Process.Start(file);
            this.Close();
        }
        private Dictionary<String, String> getTexts(String lang)
        {
            String[] lines;
            String fileName = String.Format(@"Resources\texts_{0}.txt", lang);
            lines = System.IO.File.ReadAllLines(fileName,  Encoding.Default);
            Dictionary<String, String> map = new Dictionary<String, String>();
            foreach (String line in lines)
            {
                String[] atoms= line.Split(';');
                map[atoms[0]] = atoms[1];

            }
            map["outputFolder"] = this.textBoxOutputFolder.Text;
            map["danceName"] = this.textBoxDanceName.Text;
            map["callerName"] = this.callerName;
            map["callerPictureFile"]= this.callerPictureFile;
            return map;
        }

        private void comboBoxCaller_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.callerName= comboBoxCaller.SelectedItem.ToString();
            this.callerPictureFile = Caller_dictionary[callerName];
            this.textBoxCallerPicture.Text = callerPictureFile;
            if (callerName != "" && callerPictureFile != "")
            {
                this.buttonOk.Enabled = true;
            }
        }

        private void dateTimePickerStart_ValueChanged(object sender, EventArgs e)
        {
            this.setEndDate();
        }

        private void buttonSelectOutputFolder_Click(object sender, EventArgs e)
        {

            DialogResult result = folderBrowserDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                this.textBoxOutputFolder.Text = folderBrowserDialog1.SelectedPath;
            }
        }

        private void radioButtonFestival_CheckedChanged(object sender, EventArgs e)
        {
            this.setEndDate();
        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
