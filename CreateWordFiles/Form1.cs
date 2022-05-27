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
        String callerName = "";
        String callerPictureFile = "";


        public Form1()
        {
            InitializeComponent();
            this.myInit();
        }

        #region ------------------------------------------------------------------------- Methods

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

        private static Fees getFees()
        {
            String url = Utility.map["dancing_fees"];
            String feesJson = getResponseText(url);
            Fees fees = JsonConvert.DeserializeObject<Fees>(feesJson);
            return fees;
        }

        /// <summary>
        /// Fetches data (text) via a WebRequest.
        /// Appends a random querystring to force fresh fetch (avoid cached data)
        /// </summary>
        /// <param name="url"></param>
        /// <returns></returns>
        private static string getResponseText(string url)
        {
            url = url + "?" + createRandomString(8);
            WebRequest request = HttpWebRequest.Create(url);
            WebResponse response = request.GetResponse();
            StreamReader reader = new StreamReader(response.GetResponseStream());

            string responseText = reader.ReadToEnd();
            return responseText;
        }

        private void adjustCoffee()
        {
            this.radioButtonCoffeeYes.Checked = isFestival();
            this.radioButtonCoffeeNo.Checked = !isFestival();
        }

        private void adjustEndDate()
        {
            TimeSpan duration = new TimeSpan(1 * 24, 0, 0);
            if (isFestival())
            {
                duration = new TimeSpan(3 * 24, 0, 0);
            }
            this.dateTimePickerEnd.Value = this.dateTimePickerStart.Value + duration;
        }
        private void getCallers()
        {
            string callerList = getResponseText(Utility.map["get_caller_list"]);

            String[] callerPictureFiles = callerList.Split(new char[] { ';' });
            foreach (var callerPictureFile in callerPictureFiles)
            {
                if (callerPictureFile.Length > 4 && callerPictureFile != "_spmedia_thumbs")
                {
                    String callerPictureUrl = Utility.map["caller_pictures_root"] + callerPictureFile;
                    String callerName = callerPictureFile.Substring(0, callerPictureFile.Length - 4);
                    this.comboBoxCaller.Items.Add(callerName);

                    Utility.callerDictionary.Add(callerName, callerPictureUrl);
                }
            }
            this.comboBoxCaller.Sorted = true;
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
            this.comboBoxDanceSchema.Sorted = true;
        }
        private SchemaInfo getSchemaInfo()
        {
            List<DancePass> dancpasses = new List<DancePass>();
            String schema = this.comboBoxDanceSchema.Text;
            String url = String.Format("{0}/{1}.json", Utility.map["dance_schemas"], schema);
            String schemaJson = getResponseText(url);
            SchemaInfo schemaInfo = JsonConvert.DeserializeObject<SchemaInfo>(schemaJson);
            dancpasses = schemaInfo.danceSchema;
            return schemaInfo;
        }

        private List<String> getTexts(String lang, Boolean flyer)
        {
            String[] lines;
            List<String> festivalFeesTexts = new List<String>();
            String fileName = String.Format(@"Resources\texts_{0}.csv", lang);
            lines = System.IO.File.ReadAllLines(fileName, Encoding.Default);
            //Dictionary<String, String> map = new Dictionary<String, String>();

            foreach (String line in lines)
            {
                String[] atoms = line.Split(';');
                if (atoms[0] == "festival_fees_text")
                {
                    if (atoms[3].StartsWith("---"))
                    {
                        if (flyer)
                        {
                            festivalFeesTexts.Add(";0;");
                        }
                        else
                        {
                            festivalFeesTexts.Add(";0;-");
                        }
                        atoms[3] = atoms[3].Substring(3);
                    }
                    festivalFeesTexts.Add(atoms[1] + ";" + atoms[2] + ";" + atoms[3]);
                }
                else
                {
                    Utility.map[atoms[0]] = atoms[1];
                }

            }
            Utility.map["outputFolder"] = this.textBoxOutputFolder.Text;
            Utility.map["danceName"] = this.comboBoxDanceName.Text;
            Utility.map["callerName"] = this.callerName;
            Utility.map["callerPictureFile"] = this.callerPictureFile;
            return festivalFeesTexts;
        }

        private Boolean isFestival()
        {
            return this.comboBoxDanceSchema.Text.StartsWith("festival");
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
            getCallers();
            getDanceSchemas();
            try
            {
                this.comboBoxCaller.SelectedIndex = Properties.Settings.Default.Selected_caller;
                this.comboBoxDanceSchema.SelectedIndex = Properties.Settings.Default.Selected_schema;
                this.comboBoxDanceName.SelectedIndex = Properties.Settings.Default.Selected_dance;
            }catch (Exception)
            {

            }

            int days = 2;
            if (isFestival())
            {
                days = 4;
            }
            this.dateTimePickerEnd.Value = this.dateTimePickerStart.Value + new TimeSpan((days - 1) * 24, 0, 0);
            textBoxExtra.Text = Properties.Settings.Default.Extra_se;
        }
        private void prepareForDocument(out string lang, out List<string> festivalFeesText, out SchemaInfo schemaInfo, out Fees fees, out string schemaName, out string path, Boolean flyer)
        {
            var radioButtonLanguage = groupBoxLanguage.Controls.OfType<RadioButton>()
                   .Where(r => r.Checked).FirstOrDefault();
            lang = (String)radioButtonLanguage.Tag;
            festivalFeesText = this.getTexts(lang, flyer);
            schemaInfo = this.getSchemaInfo();
            fees = getFees();
            schemaName = this.comboBoxDanceSchema.Text;
#if DEBUG
            String fileName = String.Format("{0}_{1}_{2}_{3}.docx", this.dateTimePickerStart.Value.ToShortDateString(), lang, schemaName, Utility.map["danceName"]);
            path = Path.Combine(Utility.map["outputFolder"] + "/test", fileName);
#else
            String fileName = String.Format("{0}_{1}_{2}.docx", this.dateTimePickerStart.Value.ToShortDateString(), lang, Utility.map["danceName"]);
            path = Path.Combine(Utility.map["outputFolder"], fileName);
#endif
            //try
            //{
            Properties.Settings.Default.Extra_se = textBoxExtra.Text;
            Properties.Settings.Default.Selected_caller = comboBoxCaller.SelectedIndex;
            Properties.Settings.Default.Selected_schema = comboBoxDanceSchema.SelectedIndex;
            Properties.Settings.Default.Selected_dance = this.comboBoxDanceName.SelectedIndex;

            Properties.Settings.Default.Save();
        }

        #endregion ---------------------------------------------------------------------- Methods

        #region ------------------------------------------------------------------------- Event handlers

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void buttonFlyer_Click(object sender, EventArgs e)
        {
            string lang, schemaName, path;
            List<string> festivalFeesText;
            SchemaInfo schemaInfo;
            Fees fees;
            try
            {

                prepareForDocument(out lang, out festivalFeesText, out schemaInfo, out fees, out schemaName, out path, true);

                FlyerCreator.CreateWordprocessingDocument(Utility.map, lang, schemaInfo, schemaName, path, fees, radioButtonCoffeeYes.Checked,
               (String)this.comboBoxDanceLocation.SelectedValue, this.dateTimePickerStart.Value, this.dateTimePickerEnd.Value,
               textBoxExtra.Text, festivalFeesText);
                System.Diagnostics.Process.Start(path);
            }
            catch (Exception)
            {
                MessageBox.Show(String.Format("Något gick fel."));
            }

        }

        private void buttonHtml_Click(object sender, EventArgs e)
        {
            string lang, schemaName, path;
            List<string> festivalFeesText;
            SchemaInfo schemaInfo;
            Fees fees;
            try
            {
                prepareForDocument(out lang, out festivalFeesText, out schemaInfo, out fees, out schemaName, out path, false);

                HtmlCreator.CreateHtml(Utility.map, lang, schemaInfo, path, fees, radioButtonCoffeeYes.Checked,
               (String)this.comboBoxDanceLocation.SelectedValue, this.dateTimePickerStart.Value, this.dateTimePickerEnd.Value,
               textBoxExtra.Text, festivalFeesText);
                MessageBox.Show("Klar!");


            }
            catch (Exception)
            {
                MessageBox.Show(String.Format("Något gick fel."));
            }

        }

        private void buttonSelectOutputFolder_Click(object sender, EventArgs e)
        {

            DialogResult result = folderBrowserDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                this.textBoxOutputFolder.Text = folderBrowserDialog1.SelectedPath;
            }
        }

        private void comboBoxCaller_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.callerName = comboBoxCaller.SelectedItem.ToString();
            String[] names = callerName.Split('_');
            if (names.Length >= 2)
            {
                this.callerPictureFile = Utility.callerDictionary[callerName];
                this.textBoxCallerPicture.Text = callerPictureFile;
                if (callerName != "" && callerPictureFile != "")
                {
                    this.buttonHtml.Enabled = true;
                }
            }
            else
            {
                MessageBox.Show("Fel format på callernamn. Skall vara förnamn_efternamn");
            }
        }

        private void comboBoxDanceSchema_SelectedIndexChanged(object sender, EventArgs e)
        {
            adjustEndDate();
            adjustCoffee();
        }

        private void dateTimePickerStart_ValueChanged(object sender, EventArgs e)
        {
            this.adjustEndDate();
        }
        private void textBoxCallerPicture_TextChanged(object sender, EventArgs e)
        {

        }
        #endregion ---------------------------------------------------------------------- Event handlers

        private void buttonTest_Click(object sender, EventArgs e)
        {
            String fileName = "test.docx";
            String path = Path.Combine(@"D:\Tmp", fileName);

            //GeneratedCode.GeneratedClass x = new GeneratedCode.GeneratedClass();
            //x.CreatePackage(path);
        }
    }
}
