using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CreateWordFiles
{
    public class SchemaInfo
    {
        public String schemaName { get; set; }
        public List<DancePass> danceSchema { get; set; }
        public List<int> colWidth { get; set; }
    }
    public class Fees
    {
        public List<int> weekends { get; set; }
        public List<int> festival { get; set; }
        public List<int> festival_member { get; set; }
    }
    public class DancePass
    {
        public int pass_no { get; set; }
        public int day { get; set; }
        public String start_time { get; set; }
        public String end_time { get; set; }
        public String level { get; set; }
    }
    public class Motiv8Urls
    {
        public String dance_schemas { get; set; }
        public String caller_pictures_root { get; set; }
    }

    public class Utility
    {
        /// <summary>
        /// Dictionary holding urls and texts in different languages
        /// </summary>
        public static Dictionary<String, String> map = new Dictionary<String, String>();
        public static Dictionary<String, String> danceLocations = new Dictionary<String, String>();
        public static Dictionary<string, string> callerDictionary = new Dictionary<string, string>();
        public static List<DancePass[]> dancePassesDayList = new List<DancePass[]>();
        /// <summary>
        /// Rads a semikolon separated file and adds data to the map Dictionary
        /// </summary>
        /// <param name="fileName"></param>
        public static void readToDictionary(String fileName, Dictionary<String, String> dictionary)
        {
            String[] lines = System.IO.File.ReadAllLines(fileName, Encoding.Default);
            foreach (String line in lines)
            {
                String[] atoms = line.Split(';');
                dictionary[atoms[0]] = atoms[1];
            }

        }

        public static void createFestivalRow(DancePass dancePass, int dayNumber, int passNumber, out string weekDay, out string timeString, out string level)
        {

            level = "";
            timeString = "";
            String[] weekDays = { "Fredag", "Lördag", "Söndag", "Måndag", "Tisdag" };
            weekDay = weekDays[dayNumber];
            try
            {
                timeString = formatTimeInterval(dancePass);
                level = dancePass.level;
            }
            catch (Exception)
            {

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

        public static void findBullets(List<string> festivalFeeTexts, out int bulletStart, out int bulletLength)
        {
            bulletStart = -1;
            bulletLength = 0;
            Boolean bullet = false;
            for (int i1 = 0; i1 < festivalFeeTexts.Count; i1++)
            {
                String text = festivalFeeTexts[i1];
                String[] atoms = text.Split(';');
                if (atoms[0] == "b" && !bullet)
                {
                    bulletStart = i1;
                    bulletLength = 1;
                    bullet = true;
                }
                else if (atoms[0] == "b")
                {
                    bulletLength++;
                }
            }
        }
        public static void generateFestivalFeeLines(DateTime danceDateStart, List<string> festivalFeeTexts, Fees myFees, List<string> lines)
        {
            for (int i1 = 0; i1 < festivalFeeTexts.Count; i1++)
            {
                String text = festivalFeeTexts[i1];

                String[] atoms = text.Split(';');
                DateTime dueDate1 = danceDateStart - new TimeSpan(4 * 24, 0, 0);
                String ddMM1 = dueDate1.Day + "/" + dueDate1.Month;
                int n = Int32.Parse(atoms[1]);
                switch (n)
                {
                    case 1:
                        String txt1 = String.Format(atoms[2], myFees.festival[0], myFees.festival[1],
                            myFees.festival[2], myFees.festival[3],
                            myFees.festival[4], myFees.festival[5]);
                        lines.Add(txt1);
                        break;
                    case 2:
                        String txt2 = String.Format(atoms[2], ddMM1);
                        lines.Add(txt2);
                        break;
                    case 3:
                        String txt3 = String.Format(atoms[2], myFees.festival_member[0], myFees.festival_member[1],
                            myFees.festival_member[2], myFees.festival_member[3],
                            myFees.festival_member[4], myFees.festival_member[5]);
                        lines.Add(txt3);
                        break;
                    case 4:
                        String txt4 = String.Format(atoms[2], ddMM1);
                        lines.Add(txt4);
                        break;
                    case 5:
                        String txt5 = String.Format(atoms[2], danceDateStart.Day + "/" + danceDateStart.Month);
                        lines.Add(txt5);
                        break;
                    default:
                        lines.Add(atoms[2]);
                        break;
                }


            }
        }
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

        public static DancePass[] getDancePassesForDay(List<DancePass> dancePasses, int day)
        {
            IEnumerable<DancePass> dayOnePasses = from dancePass in dancePasses
                                                  where dancePass.day == day
                                                  orderby dancePass.pass_no ascending
                                                  select dancePass;
            var dayOnePassesArray = dayOnePasses.ToArray();
            return dayOnePassesArray;
        }


    }

}
