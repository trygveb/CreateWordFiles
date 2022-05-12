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
        //private String callerPicturesRoot;
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
            Utility.readToDictionary(@"Resources\dance_locations.csv", Utility.danceLocations);
            comboBoxDanceLocation.DataSource = new BindingSource(Utility.danceLocations, null);
            comboBoxDanceLocation.DisplayMember = "Key";
            comboBoxDanceLocation.ValueMember = "Value";
            //foreach (KeyValuePair<string, string> location in Utility.danceLocations)
            //{
            //    comboBoxDanceLocation.Items.Add(location);
            //}
            getCallers();
            getDanceSchemas();
            this.comboBoxCaller.SelectedIndex = 0;
            this.comboBoxDanceSchema.SelectedIndex = 4; 
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

        private SchemaInfo getSchemaInfo()
        {
            List<DancePass> dancpasses = new List<DancePass>();
            String schema= this.comboBoxDanceSchema.Text;
            String url = String.Format("{0}/{1}.json", Utility.map["dance_schemas"], schema);
            String schemaJson = getResponseText(url);
            SchemaInfo schemaInfo= JsonConvert.DeserializeObject<SchemaInfo > (schemaJson);
            dancpasses = schemaInfo.danceSchema;
            return schemaInfo;
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
        /// <summary>
        /// Create a random (letter) string of given length
        /// </summary>
        /// <param name="length"></param>
        /// <returns></returns>
        private static String createRandomString(int length = 7)
        {

            StringBuilder str_build = new StringBuilder();
            Random random = new Random();

            char letter;

            for (int i = 0; i < length; i++)
            {
                double flt = random.NextDouble();
                int shift = Convert.ToInt32(Math.Floor(25 * flt));
                letter = Convert.ToChar(shift + 65);
                str_build.Append(letter);
            }
            return (str_build.ToString());
        }

        /// <summary>
        /// Fetches data (text) via a WebRequest.
        /// Appends a random querystring to force fresh fetch (avoid cached data)
        /// </summary>
        /// <param name="url"></param>
        /// <returns></returns>
        private static string getResponseText(string url)
        {
            url = url + "?"+ createRandomString(8);
            WebRequest request = HttpWebRequest.Create(url);
            WebResponse response = request.GetResponse();
            StreamReader reader = new StreamReader(response.GetResponseStream());

            string responseText = reader.ReadToEnd();
            return responseText;
        }
        private void adjustEndDate()
        {
            TimeSpan duration = new TimeSpan(1 * 24, 0, 0);
            if (this.comboBoxDanceSchema.Text == "festival")
            {
                duration = new TimeSpan(3 * 24, 0, 0);
            }
            this.dateTimePickerEnd.Value = this.dateTimePickerStart.Value + duration;
        }

        private void buttonOk_Click(object sender, EventArgs e)
        {
            var radioButtonLanguage = groupBoxLanguage.Controls.OfType<RadioButton>()
                   .Where(r => r.Checked).FirstOrDefault();
            String lang = (String)radioButtonLanguage.Tag;
            this.getTexts(lang);
            SchemaInfo schemaInfo = this.getSchemaInfo();

            Fees fees= getFees();

            String schemaName = this.comboBoxDanceSchema.Text;
#if DEBUG
            String fileName = String.Format("{0}_{1}_{2}_{3}.docx", this.dateTimePickerStart.Value.ToShortDateString(), lang, schemaName, Utility.map["danceName"] );
            String path = Path.Combine(Utility.map["outputFolder"]+"/test", fileName);
#else
            String fileName = String.Format("{0}_{1}_{2}.docx", this.dateTimePickerStart.Value.ToShortDateString(), lang, Utility.map["danceName"]);
            String path = Path.Combine(Utility.map["outputFolder"], fileName);
#endif
            try
            {
                Creator.CreateWordprocessingDocument(Utility.map, lang, schemaInfo, schemaName, path, fees, radioButtonCoffeeYes.Checked,
                (String) this.comboBoxDanceLocation.SelectedValue, this.dateTimePickerStart.Value, this.dateTimePickerEnd.Value);
                System.Diagnostics.Process.Start(path);
            }
            catch (Exception)
            {
                MessageBox.Show(String.Format("Något gick fel. Är filen {0} öppen?", path));
            }

        }

        private static Fees getFees()
        {
            String url = Utility.map["dancing_fees"];
            String feesJson = getResponseText(url);
            Fees fees = JsonConvert.DeserializeObject<Fees>(feesJson);
            return fees;
        }

        private void getTexts(String lang)
        {
            String[] lines;
            String fileName = String.Format(@"Resources\texts_{0}.csv", lang);
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
            this.adjustEndDate();
        }

        private void buttonSelectOutputFolder_Click(object sender, EventArgs e)
        {

            DialogResult result = folderBrowserDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                this.textBoxOutputFolder.Text = folderBrowserDialog1.SelectedPath;
            }
        }


        private void buttonCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

 
        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void textBoxCallerPicture_TextChanged(object sender, EventArgs e)
        {

        }

        private void comboBoxDanceSchema_SelectedIndexChanged(object sender, EventArgs e)
        {
            adjustEndDate();
        }

    }
}
