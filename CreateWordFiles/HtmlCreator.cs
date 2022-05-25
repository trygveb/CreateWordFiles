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
        public static StringBuilder htmlStringBuilder = new StringBuilder();
//        private static List<DancePass[]> dancePassesDayList = new List<DancePass[]>();
        private static Fees myFees;
        private static Dictionary<String, String> myTexts;
        public static void CreateDanceSchemaHtmlTable(String lang, SchemaInfo schemaInfo, int fontSize, int lineSpace)
        {
            int numberOfDistinctDays = Utility.createDancePassesDaylist(schemaInfo);

            if (numberOfDistinctDays == 2)
            {
                htmlStringBuilder.Append("<table class='m8_schema'>");
                createWeekendDanceSchemaTable(lang, Utility.dancePassesDayList, schemaInfo, fontSize, lineSpace);
                htmlStringBuilder.Append("</table>");
            }
            else if (numberOfDistinctDays == 4)
            {
                CreateFestivalDanceSchemaHtmlTable(lang, Utility.dancePassesDayList, schemaInfo);
                htmlStringBuilder.Append("</table>");
            }

        }

        public static void CreateFestivalDanceSchemaHtmlTable(String lang, List<DancePass[]> dancePassesDayList, SchemaInfo schemaInfo)
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

        public static void CreateHtml(Dictionary<String, String> texts, String lang, SchemaInfo schemaInfo,
                            String path, Fees fees, Boolean coffee, String danceLocation, DateTime danceDateStart,
           DateTime danceDateEnd, String extra, List<String> festivalFeesText)
        {
            myTexts = texts;
            myFees = fees;
            if (schemaInfo.schemaName.Contains("festival"))
            {
                createFestivalHtml(lang, schemaInfo, coffee, danceLocation, danceDateStart, festivalFeesText);

            }
            else
            {
                createWeekendHtml(lang, schemaInfo, coffee, danceLocation, danceDateStart);
            }
            String htmlText = htmlStringBuilder.ToString();
            File.WriteAllText(path.Replace("docx", "htm"), htmlText);

            createDemoHtmlFile(path, htmlText);
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

        private static void createFestivalHtml(string lang, SchemaInfo schemaInfo, bool coffee, string danceLocation, DateTime danceDateStart, List<string> festivalFeesText)
        {
            htmlStringBuilder.Clear();
            htmlStringBuilder.Append(String.Format("<p class='m8_schema m8_border' style='max-width: 450px;'>{0}</p><br>", danceLocation));

            htmlStringBuilder.Append(String.Format("<p class='m8'>\n"));
            htmlStringBuilder.Append(String.Format("<span style='font-size:larger; font-weight:bold; text-decoration:underline;'>PROGRAM</span>\n"));
            htmlStringBuilder.Append(String.Format("</p>\n"));
            CreateDanceSchemaHtmlTable(lang, schemaInfo, 12, 200);
            htmlStringBuilder.Append(String.Format("<p class='m8'>{0}</p><br>\n", myTexts["rotation"]));

            htmlStringBuilder.Append(String.Format("<p class='m8'>\n"));
            htmlStringBuilder.Append(String.Format("<span style='font-size:larger; font-weight:bold; text-decoration:underline;'>ENTRÉ</span>\n"));
            htmlStringBuilder.Append(String.Format("</p>\n"));

            generateFestivalFeesHtml(schemaInfo, danceDateStart, festivalFeesText);

            String text1 = myTexts["no_coffee"];
            if (coffee)
            {
                text1 = myTexts["coffee"];
            }
            htmlStringBuilder.Append(String.Format("<br><p class='m8_schema m8_border' style='max-width: 450px;'>{0}</p><br>",
                text1 + " - " + myTexts["lunch"]));
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
        private static void createWeekendDanceSchemaTable(String lang, List<DancePass[]> dancePassesDayList, SchemaInfo schemaInfo,
            int fontSize, int lineSpace)
        {

            htmlStringBuilder.Append("<tr class='m8_schema'>" +
               "<th colspan=2 class='m8_schema'>" + myTexts["Saturday"] + "</th>" +
               "<th class='m8_schema m8_space' style='min-width:50px;'></th>" +
                "<th colspan=2 class='m8_schema'>" + myTexts["Sunday"] + "</th>" +
                "</tr>\n");

            htmlStringBuilder.Append(createWeekEndRowHtml(lang, dancePassesDayList, 2, schemaInfo));

            htmlStringBuilder.Append(createWeekEndRowHtml(lang, dancePassesDayList, 3, schemaInfo, schemaInfo.schemaName == "weekend_meeting"));

            if (dancePassesDayList[0].Length > 2 || dancePassesDayList[1].Length > 2)
            {
                htmlStringBuilder.Append(createWeekEndRowHtml(lang, dancePassesDayList, 4, schemaInfo));
            }

            //  return table;

        }

        private static void createWeekendHtml(string lang, SchemaInfo schemaInfo, bool coffee, string danceLocation, DateTime danceDateStart)
        {
            CreateDanceSchemaHtmlTable(lang, schemaInfo, 12, 200);
            generateWeekendFeesLines(schemaInfo);
        }
        private static String createWeekEndRowHtml(String lang, List<DancePass[]> dancePassesDayList, int rowNumber, SchemaInfo schemaInfo, Boolean merge = false)
        {
            string level2, timeString2, level1, timeString1;
            Utility.createWeekEndRow(dancePassesDayList, rowNumber, out level2, out timeString2, out level1, out timeString1);


            String row = String.Format("<tr class='m8_schema'><td class='m8_schema m8_time'>{0}</td><td class='m8_schema m8_level'>{1}</td><td class='m8_schema m8_space'> </td><td class='m8_schema m8_time'>{2}</td><td class='m8_schema m8_level'>{3}</td></tr>",
                timeString1, level1, timeString2, level2);
            return row;

        }

        private static void generateFestivalFeesHtml(SchemaInfo schemaInfo, DateTime danceDateStart, List<string> festivalFeeTexts)
        {
            List<String> lines = new List<String>();

            // festivalFeeTexts may have zero or one bullet lists
            //int bulletStart = -1, bulletLength = 0;
            List<string> festivalFeeTextsWithBullets, festivalFeeTextsWithoutBullets;
            Utility.findBullets(festivalFeeTexts, out festivalFeeTextsWithBullets, out festivalFeeTextsWithoutBullets);
            Utility.generateFestivalFeeLines(danceDateStart, festivalFeeTexts, myFees,  lines);
           
            for (int i1 = 0; i1 < festivalFeeTextsWithoutBullets.Count; i1++)
            {
                String line = festivalFeeTextsWithoutBullets[i1];
                htmlStringBuilder.Append(String.Format("<p class='m8'>{0}</p>\n", line));
            }
            if (festivalFeeTextsWithBullets.Count > 0)
            {
                htmlStringBuilder.Append(String.Format("<ul  class='m8'>\n"));
                for (int i1 = 0; i1 < festivalFeeTextsWithBullets.Count; i1++)
                {
                    String line = festivalFeeTextsWithBullets[i1];
                    htmlStringBuilder.Append(String.Format("<li class='m8'>{0}</li>\n", line));
                }
                htmlStringBuilder.Append(String.Format("</ul>\n"));
            }
        }
        private static void generateWeekendFeesLines(SchemaInfo schemaInfo)
        {

            List<String> lines = new List<String>();
            if (schemaInfo.schemaName.StartsWith("weekend"))
            {
                String line1 = String.Format(myTexts["weekend_member_fees"], myFees.weekends[0], myFees.weekends[1]);
                htmlStringBuilder.Append("<p class='m8_schema'>\n");
                htmlStringBuilder.Append(line1 + "<br>");
                String line2 = String.Format(myTexts["weekend_non_member_fees"], myFees.weekends[2], myFees.weekends[3]);
                htmlStringBuilder.Append(line2 + "<br>");
                if (schemaInfo.schemaName == "weekend_january")
                {
                    htmlStringBuilder.Append(myTexts["one_pass_sunday"] + "<br>");
                }
                String pgPay = myTexts["weekend_pg_pay"];
                if (pgPay != "N/A")
                {
                    htmlStringBuilder.Append(pgPay + "<br>");
                }
                String swishpay = myTexts["weekend_swish_pay"];
                if (swishpay != "N/A")
                {
                    htmlStringBuilder.Append(swishpay + "<br>");
                }
            }
            htmlStringBuilder.Append("</p>");
        }
    }
}
