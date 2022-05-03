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
       public partial class Form1 : Form
    {
        private String callerPicturesRoot;
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
            Utility.readToDictionary(@"Resources\web.csv", Utility.map);

            getCallers();
            getDanceSchemas();
            this.comboBoxCaller.SelectedIndex = 0;
            this.comboBoxDanceSchema.SelectedIndex = 0; 
            this.dateTimePickerEnd.Value = this.dateTimePickerStart.Value + new TimeSpan(24, 0, 0);
        }

        private void getCallers()
        {
            string callerList = getResponseText(Utility.map["get_caller_list"]);

            String[] callerPictureFiles = callerList.Split(new char[] { ';' });
            foreach (var callerPictureFile in callerPictureFiles)
            {
                if (callerPictureFile.Length > 4)
                {
                    String callerPictureUrl = Utility.map["caller_pictures_root"] + callerPictureFile;
                    String callerName = callerPictureFile.Substring(0, callerPictureFile.Length - 4);
                    this.comboBoxCaller.Items.Add(callerName);

                    Utility.callerDictionary.Add(callerName, callerPictureUrl);
                }
            }
        }

        private List<DancePass> getDanceSchema()
        {
            List<DancePass> dancpasses = new List<DancePass>();
            String schema= this.comboBoxDanceSchema.Text;
            String url = String.Format("{0}/{1}.json", Utility.map["dance_schemas"], schema);
            String schemaJson = getResponseText(url);
            dancpasses = JsonConvert.DeserializeObject<List<DancePass>>(schemaJson);
            return dancpasses;
        }


        private void getDanceSchemas()
        {
            string schemaList = getResponseText(Utility.map["get_schema_list"]);

            String[] schemas = schemaList.Split(new char[] { ';' });
            foreach (var schema in schemas)
            {
                if (schema.Length > 5)
                {
                    this.comboBoxDanceSchema.Items.Add(schema.Substring(0, schema.Length - 5));
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
            this.dateTimePickerEnd.Value = this.dateTimePickerStart.Value + ts;
        }
        private void buttonOk_Click(object sender, EventArgs e)
        {
            var radioButtonLanguage = groupBoxLanguage.Controls.OfType<RadioButton>()
                   .Where(r => r.Checked).FirstOrDefault();
            String lang = (String)radioButtonLanguage.Tag;
            this.getTexts(lang);
            List<DancePass> danceSchema = this.getDanceSchema();
            Creator.CreateWordprocessingDocument(Utility.map, lang, danceSchema, this.dateTimePickerStart.Value, this.dateTimePickerEnd.Value);
            //System.Diagnostics.Process.Start(file);
            MessageBox.Show("Flyer skapad");
            //this.Close();
        }
        private void getTexts(String lang)
        {
            String[] lines;
            String fileName = String.Format(@"Resources\texts_{0}.txt", lang);
            lines = System.IO.File.ReadAllLines(fileName,  Encoding.Default);
            //Dictionary<String, String> map = new Dictionary<String, String>();
            foreach (String line in lines)
            {
                String[] atoms= line.Split(';');
                Utility.map[atoms[0]] = atoms[1];

            }
            Utility.map["outputFolder"] = this.textBoxOutputFolder.Text;
            Utility.map["danceName"] = this.textBoxDanceName.Text;
            Utility.map["callerName"] = this.callerName;
            Utility.map["callerPictureFile"]= this.callerPictureFile;
        }

        private void comboBoxCaller_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.callerName= comboBoxCaller.SelectedItem.ToString();
            String[] names = callerName.Split('_');
            if (names.Length == 2)
            {
                this.callerPictureFile = Utility.callerDictionary[callerName];
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
