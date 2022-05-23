using System;
using System.IO;
using System.Drawing;
using System.Globalization;
using System.Collections.Generic;
using System.Net.Cache;
using System.Linq;
using System.Text;

namespace CreateWordFiles
{
    internal class HtmlCreator
    {
         private static Dictionary<String, String> myTexts;
        private static List<DancePass[]> dancePassesDayList = new List<DancePass[]>();
        private static Fees myFees;
        public static StringBuilder htmlStringBuilder = new StringBuilder();

        public static void createHtml(Dictionary<String, String> texts, String lang, SchemaInfo schemaInfo,
           String schemaName, String path, Fees fees, Boolean coffee, String danceLocation, DateTime danceDateStart,
           DateTime danceDateEnd, String extra, List<String> festivalFeesText)
        {
            myTexts = texts;
            myFees = fees;

            htmlStringBuilder.Clear();
            htmlStringBuilder.Append(String.Format("<p class='m8_schema m8_border' style='max-width: 450px;'>{0}</p><br>", danceLocation));

            htmlStringBuilder.Append(String.Format("<p class='m8'>\n"));
            htmlStringBuilder.Append(String.Format("<span style='font-size:larger; font-weight:bold; text-decoration:underline;'>PROGRAM</span>\n"));
            htmlStringBuilder.Append(String.Format("</p>\n"));
            CreateDanceSchemaHtmlTable(lang, schemaInfo, schemaName, 12, 200);
            htmlStringBuilder.Append(String.Format("<p class='m8'>{0}</p><br>\n", myTexts["rotation"]));

            htmlStringBuilder.Append(String.Format("<p class='m8'>\n"));
            htmlStringBuilder.Append(String.Format("<span style='font-size:larger; font-weight:bold; text-decoration:underline;'>ENTRÉ</span>\n"));
            htmlStringBuilder.Append(String.Format("</p>\n"));

            GenerateFestivalFeesHtml(schemaInfo, danceDateStart, festivalFeesText);

            String text1 = myTexts["no_coffee"];
            if (coffee)
            {
                text1 = myTexts["coffee"];
            }
            htmlStringBuilder.Append(String.Format("<br><p class='m8_schema m8_border' style='max-width: 450px;'>{0}</p><br>",
                text1 + " - " + myTexts["lunch"]));


            String htmlText = htmlStringBuilder.ToString();
            File.WriteAllText(path.Replace("docx", "htm"), htmlText);

            createDemoHtmlFile(path, htmlText);
        }
        public static void CreateDanceSchemaHtmlTable(String lang, SchemaInfo schemaInfo, String schemaName, int fontSize, int lineSpace)
        {
            int numberOfDistinctDays = Utility.createDancePassesDaylist(schemaInfo);

            if (numberOfDistinctDays == 2)
            {
                htmlStringBuilder.Append("<table class='m8_schema'>");
                // Wp.Table table = createWeekendDanceSchemaTable(lang, dancePassesDayList, schemaInfo, fontSize, lineSpace);
                htmlStringBuilder.Append("</table>");
            }
            else if (numberOfDistinctDays == 4)
            {
                createFestivalDanceSchemaHtmlTable(lang, dancePassesDayList, schemaInfo);
                htmlStringBuilder.Append("</table>");
            }

        }
        private static void GenerateFestivalFeesHtml(SchemaInfo schemaInfo, DateTime danceDateStart, List<string> festivalFeeTexts)
        {
            List<String> lines = new List<String>();

            // festivalFeeTexts may have zero or one bullet lists
            int bulletStart = -1, bulletLength = 0;
            Utility.findBullets(festivalFeeTexts, out bulletStart, out bulletLength);
            Utility.generateFestivalFeeLines(danceDateStart, festivalFeeTexts, myFees,  lines);
            Boolean bullet = false;
            for (int i1 = 0; i1 < lines.Count; i1++)
            {
                String line = lines[i1];
                if (i1 == bulletStart)
                {
                    htmlStringBuilder.Append(String.Format("<ul  class='m8'>\n"));
                    bullet = true;
                }
                if (bullet)
                {
                    htmlStringBuilder.Append(String.Format("<li class='m8'>{0}</li>\n", line));
                    if (i1 == bulletStart + bulletLength)
                    {
                        bullet = false;
                        htmlStringBuilder.Append(String.Format("</ul>\n", line));
                    }
                }
                else
                {
                    htmlStringBuilder.Append(String.Format("<p class='m8'>{0}</p>\n", line));
                }
            }
            htmlStringBuilder.Append(String.Format("</ul>\n"));
        }
        /// <summary>
        /// Creates a demo html file with html header and style tag
        /// </summary>
        /// <param name="path"></param>
        /// <param name="htmlText"></param>
        private static void createDemoHtmlFile(string path, string htmlText)
        {
            String css = File.ReadAllText(@"Resources\motiv8.css");
            String yyy = String.Format(@"<!DOCTYPE html>
                    <html>
                    <head>
                        <style>
                        {0}
                        </style >
                    </head >
                    <body> ", css);
            File.WriteAllText(path.Replace("docx", "html"), yyy + htmlText + "</body></html>");
        }

        public static void createFestivalDanceSchemaHtmlTable(String lang, List<DancePass[]> dancePassesDayList, SchemaInfo schemaInfo)
        {
            htmlStringBuilder.Append("<table class='m8_festival'>");

            int dayNumber = 0;
            int passNumber = 1;

            foreach (DancePass[] dancePasses in dancePassesDayList)
            {
                foreach (DancePass dancePass in dancePasses)
                {
                    createFestivalProgramRowForHtml(dayNumber, passNumber, dancePass, schemaInfo.colWidth);
                    passNumber++;
                }
                dayNumber++;

            }
        }

        private static void createFestivalProgramRowForHtml(int dayNumber, int passNumber, DancePass dancePass, List<int> colWidth)
        {
            string weekDay, timeString, level;
            Utility.createFestivalRow(dancePass, dayNumber, passNumber, out weekDay, out timeString, out level);

            String[] content = { passNumber.ToString(), weekDay, timeString, level };
            htmlStringBuilder.Append("<tr class='m8_festival'>\n");
            for (int i = 0; i < content.Length; i++)
            {
                if (dayNumber == 1) // Saturday
                {
                    if (passNumber == 2 && i == 1)  // First pass on Saturday, second column should span 3 rows
                    {
                        htmlStringBuilder.Append(String.Format("<td rowspan='3' class='m8_festival'>{0}</td>", content[i]));
                    }
                    else if (i != 1)  // For other passes on Saturday, skip second column
                    {
                        htmlStringBuilder.Append(String.Format("<td class='m8_festival'>{0}</td>", content[i]));
                    }
                }
                else if (dayNumber == 2) // Sunday
                {
                    if (passNumber == 5 && i == 1)  // First pass on Sunday, second column should span 2 rows
                    {
                        htmlStringBuilder.Append(String.Format("<td rowspan='2' class='m8_festival'>{0}</td>", content[i]));
                    }
                    else if (i != 1)  // For other passes on Sunday, skip second column
                    {
                        htmlStringBuilder.Append(String.Format("<td class='m8_festival'>{0}</td>", content[i]));
                    }
                }
                else
                {
                    htmlStringBuilder.Append(String.Format("<td class='m8_festival'>{0}</td>", content[i]));
                }
            }
            htmlStringBuilder.Append("</tr>\n");

        }


    }
}
