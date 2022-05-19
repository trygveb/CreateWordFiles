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
    }
}
