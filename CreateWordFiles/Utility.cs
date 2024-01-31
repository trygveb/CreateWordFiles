using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Globalization;
using System.Threading.Tasks;

namespace CreateWordFiles
{
    public class DancePass
    {
        public int day { get; set; }
        public String end_time { get; set; }
        public String level { get; set; }
        public int pass_no { get; set; }
        public String start_time { get; set; }
    }

    public class Fees
    {
        public List<int> festival { get; set; }
        public List<int> festival_member { get; set; }
        public List<int> weekends { get; set; }
    }

    public class Motiv8Urls
    {
        public String caller_pictures_root { get; set; }
        public String dance_schemas { get; set; }
    }

    public class SchemaInfo
    {
        public List<int> colWidth { get; set; }
        public List<DancePass> danceSchema { get; set; }
        public String schemaName { get; set; }
    }
    public class Utility

    {
        public static DateTimeFormatInfo dateTimeFormatInfo_se = CultureInfo.GetCultureInfo("sv-SE").DateTimeFormat;
        public static DateTimeFormatInfo dateTimeFormatInfo_en = CultureInfo.GetCultureInfo("en-US").DateTimeFormat;

        public static Dictionary<string, string> callerDictionary = new Dictionary<string, string>();

        public static Dictionary<String, String> danceLocations = new Dictionary<String, String>();

        public static List<DancePass[]> dancePassesDayList = new List<DancePass[]>();

        /// <summary>
        /// Dictionary holding urls and texts in different languages
        /// </summary>
        public static Dictionary<String, String> map = new Dictionary<String, String>();
        public static int createDancePassesDaylist(SchemaInfo schemaInfo)
        {
            List<DancePass> dancePasses = schemaInfo.danceSchema;

            var n = dancePasses.Select(o => new { Day = o.day }).Distinct();
            int numberOfDistinctDays = n.Count();

            dancePassesDayList.Clear();

            for (int i = 1; i <= numberOfDistinctDays; i++)
            {
                dancePassesDayList.Add(getDancePassesForDay(dancePasses, i));
            }

            return numberOfDistinctDays;
        }

        public static void createFestivalRow(String lang, DancePass dancePass, int dayNumber, int passNumber, out string weekDay, out string timeString, out string level)
        {

            level = "";
            timeString = "";
            //String[] weekDays = { "Fredag", "Lördag", "Söndag", "Måndag" };
            Dictionary<String, String> weekDaysX = new Dictionary<String, String>    {
                { "en0", "Friday" },
                { "en1", "Saturday"},
                { "en2", "Sunday"},
                { "en3", "Monday"},
                { "se0", "Fredag"},
                { "se1", "Lördag"},
                { "se2", "Söndag"},
                { "se3", "Måndag"}
            };
            String key = String.Format("{0}{1}", lang, dayNumber);
            weekDay = weekDaysX[key];
            try
            {
                timeString = formatTimeInterval(dancePass);
                level = dancePass.level;
            }
            catch (Exception)
            {

            }
        }

        public static void createWeekEndRow(List<DancePass[]> dancePassesDayList, int rowNumber, out string level2, out string timeString2, out string level1, out string timeString1)
        {
            int i1 = rowNumber - 2;
            level2 = "";
            timeString2 = "";
            try
            {
                timeString2 = Utility.formatTimeInterval(dancePassesDayList[1][i1]);
                level2 = dancePassesDayList[1][i1].level;
            }
            catch (Exception)
            {

            }
            level1 = "";
            timeString1 = "";
            try
            {
                timeString1 = Utility.formatTimeInterval(dancePassesDayList[0][i1]);
                level1 = dancePassesDayList[0][i1].level;
            }
            catch (Exception)
            {

            }
        }

        public static void findBullets(List<string> festivalFeeTexts, out List<string> festivalFeeTextsWithBullets, out List<string> festivalFeeTextsWithoutBullets)
        {
            festivalFeeTextsWithBullets = new List<string>();
            festivalFeeTextsWithoutBullets = new List<string>();
            for (int i1 = 0; i1 < festivalFeeTexts.Count; i1++)
            {
                String text = festivalFeeTexts[i1];
                String[] atoms = text.Split(';');
                if (atoms[0] == "b")
                {
                    festivalFeeTextsWithBullets.Add(text);
                }
                else
                {
                    festivalFeeTextsWithoutBullets.Add(text);
                }
            }
        }

        public static String formatTimeInterval(DancePass dancePass)
        {
            String text = String.Format("{0} - {1}", dancePass.start_time, dancePass.end_time);
            if (dancePass.end_time.Length > 6) // quick and dirty solution for Årsmöte
            {
                text = String.Format("{0} {1}", dancePass.start_time, dancePass.end_time);
            }
            return text;
        }

        public static void generateFestivalFeeLines(DateTime danceDateStart, List<string> festivalFeeTexts, Fees myFees, List<string> lines)
        {
            for (int i1 = 0; i1 < festivalFeeTexts.Count; i1++)
            {
                String text = festivalFeeTexts[i1];

                String[] atoms = text.Split(';');
                DateTime dueDate1 = danceDateStart - new TimeSpan(7 * 24, 0, 0);
                String ddMM1 = dueDate1.Day + "/" + dueDate1.Month;
                int n = Int32.Parse(atoms[1]);
                if (atoms[0] == "l")
                {
                    lines.Add("");
                }
                switch (n)
                {
                    case 1: // Line with the fees for 1-6 sessions
                        String txt1 = String.Format(atoms[2], myFees.festival[0], myFees.festival[1],
                            myFees.festival[2], myFees.festival[3],
                            myFees.festival[4], myFees.festival[5]);
                        lines.Add(txt1);
                        break;
                    case 2:  // Line with due date for advanced payments
                        String txt2 = String.Format(atoms[2], ddMM1);
                        lines.Add(txt2);
                        break;
                    case 3: // Line with the MEMBER fees for 1-6 sessions
                        String txt3 = String.Format(atoms[2], myFees.festival_member[0], myFees.festival_member[1],
                            myFees.festival_member[2], myFees.festival_member[3],
                            myFees.festival_member[4], myFees.festival_member[5]);
                        lines.Add(txt3);
                        break;
                    case 4: // Line with date for first dance day
                        String txt5 = String.Format(atoms[2], danceDateStart.Day + "/" + danceDateStart.Month);
                        lines.Add(txt5);
                        break;
                    default:
                        lines.Add(atoms[2]);
                        break;
                }


            }
        }

              public static DancePass[] getDancePassesForDay(List<DancePass> dancePasses, int day)
        {
            IEnumerable<DancePass> dayOnePasses = from dancePass in dancePasses
                                                  where dancePass.day == day
                                                  orderby dancePass.pass_no ascending
                                                  select dancePass;
            var dayOnePassesArray = dayOnePasses.ToArray();
            return dayOnePassesArray;
        }

        /// <summary>
        /// Reads a semikolon separated file and adds data to the map Dictionary
        /// First CSV field becomes the key, the second becomes the value of the dictionary
        /// </summary>
        /// <param name="fileName" ></param>
        /// <param name="dictionary"></param>
        public static void readToDictionary(String fileName, Dictionary<String, String> dictionary)
        {
            String[] lines = System.IO.File.ReadAllLines(fileName, Encoding.Default);
            foreach (String line in lines)
            {
                String[] atoms = line.Split(';');
                dictionary[atoms[0]] = atoms[1];
            }

        }
        /// <summary>
        /// Reads a semikolon separated file and adds data to the dictionary
        /// First CSV field becomes the key, the second becomes the value of the dictionary, BUT
        /// if the second CSV field contains the word "_root", the value of the dictionary is 
        /// the dictionary value of that field plus the third field
        /// </summary>
        /// <param name="fileName" ></param>
        /// <param name="dictionary"></param>
        public static void readToDictionary_1(String fileName, Dictionary<String, String> dictionary)
        {
            String[] lines = System.IO.File.ReadAllLines(fileName, Encoding.Default);
            foreach (String line in lines)
            {
                String[] atoms = line.Split(';');
                if (atoms[1].Contains("_root") )
                {
                    dictionary[atoms[0]] = dictionary[atoms[1]]+ atoms[2];
                } else
                {
                    dictionary[atoms[0]] = atoms[1];
                }
                
            }

        }

    }

}
