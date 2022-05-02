using System;
using System.Collections.Generic;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Data;
using System.Net;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;

namespace CreateWordFiles
{
    public class DancePass
    {
        public int pass_no { get; set; }
        public int day { get; set; }
        public String start_time { get; set; }
        public String end_time { get; set; }
        public String level { get; set; }
    }
    class datamodel
    {
        public string key1 { get; set; }
        public string key2 { get; set; }
        public string key3 { get; set; }
    }
    public partial class Form1 : Form
    {
        private String callerPicturesRoot = "https://www.motiv8s.se/19/images/callers/";
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

            getCallers();
            this.dateTimePickerEnd.Value = this.dateTimePickerStart.Value + new TimeSpan(24, 0, 0);
        }

        private void getCallers()
        {
            var jsonFile = "https://motiv8s.se/19/danstider/jesper_dans.json";
            string jsonText = getResponseText(jsonFile);
            DancePass[] dancePasses= JsonConvert.DeserializeObject<DancePass[]>(jsonText);

            var url = "https://motiv8s.se/19/caller_list.php";
            string callerList = getResponseText(url);
            String[] callerPictureFiles = callerList.Split(new char[] { ';' });
            foreach (var callerPictureFile in callerPictureFiles)
            {
                if (callerPictureFile.Length > 4)
                {
                    String callerPictureUrl = this.callerPicturesRoot + callerPictureFile;
                    String callerName = callerPictureFile.Substring(0, callerPictureFile.Length - 4);
                    this.comboBoxCaller.Items.Add(callerName);

                    Caller_dictionary.Add(callerName, callerPictureUrl);
                }
            }


        }

        private static string getResponseText(string url)
        {
            WebRequest request = HttpWebRequest.Create(url);
            WebResponse response = request.GetResponse();
            StreamReader reader = new StreamReader(response.GetResponseStream());

            string responseText = reader.ReadToEnd();
            return responseText;
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
            MessageBox.Show("Flyer skapad");
            //this.Close();
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
            String[] names = callerName.Split('_');
            if (names.Length == 2)
            {
                this.callerPictureFile = Caller_dictionary[callerName];
                this.textBoxCallerPicture.Text = callerPictureFile;
                if (callerName != "" && callerPictureFile != "")
                {
                    this.buttonOk.Enabled = true;
                }
            }
            else
            {
                MessageBox.Show("Fel format på callernamn. Skall vara förnamn_efternamn");
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
