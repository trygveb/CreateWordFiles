using System;
using System.IO;
using System.Drawing;
using System.Globalization;
using System.Collections.Generic;
using System.Net.Cache;
using System.Linq;
using System.Text;
using OXML = DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Wp = DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace CreateWordFiles
{
    internal class Creator
    {
        /*
         CSS example

         */
        private static Wp.Color wpColorBlackx = new Wp.Color() { Val = "000000" };
        private static Wp.Color wpColorRedx = new Wp.Color() { Val = "FF0000" };
        private static Dictionary<String, String> myTexts;
        private static List<DancePass[]> dancePassesDayList = new List<DancePass[]>();
        private static Fees myFees;
        private static StringBuilder htmlStringBuilder = new StringBuilder();

        /// <summary>
        /// Only weekend dances supported 
        /// </summary>
        /// <param name="filepath"></param>
        /// <param name="danceName"></param>
        /// <param name="danceDateStart"></param>
        /// <param name="danceDateEnd"></param>
        public static void CreateWordprocessingDocument(Dictionary<String, String> texts, String lang, SchemaInfo schemaInfo,
            String schemaName, String path, Fees fees, Boolean coffee, String danceLocation, DateTime danceDateStart, DateTime danceDateEnd)
        {
            string monthName1, monthName2, danceDates;

            myTexts = texts;
            myFees = fees;

            danceDates = createDanceDates(danceDateStart, danceDateEnd, out monthName1, out monthName2);

            htmlStringBuilder.Clear();

            try
            {
                using (WordprocessingDocument wordDocument =
                    WordprocessingDocument.Create(path, OXML.WordprocessingDocumentType.Document))
                {
                    if (schemaName == "festival")
                    {
                        createFestivalFlyer(lang, schemaInfo, danceDateStart, schemaName, path, coffee, danceLocation, danceDates, wordDocument);
                    }
                    else
                    {
                        createWeekendFlyer(lang, schemaInfo, danceDateStart, schemaName, path, coffee, danceLocation, danceDates, wordDocument);
                    }
                }
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        private static void createFestivalFlyer(string lang, SchemaInfo schemaInfo, DateTime danceDateStart, string schemaName, string path, bool coffee, string danceLocation, string danceDates, WordprocessingDocument wordDocument)
        {
            MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();

            mainPart.Document = new Wp.Document();
            Wp.Body body = mainPart.Document.AppendChild(new Wp.Body());
            //MyOpenXml.SetMarginsAndFooter(wordDocument, myTexts["footer"], 10, 1500, 2000, 2000);
            MyOpenXml.SetMarginsAndFooter(wordDocument, myTexts["footer"], 10, 1000);

            Wp.Paragraph paragraph1 = GenerateWelcomeParagraphFestival1();
            body.AppendChild(paragraph1);
            addImage("Inline", wordDocument, myTexts["logo_file_name"], 200, 6.0, 10.0);
            Wp.Paragraph paragraph2 = GenerateWelcomeParagraphFestival2(myTexts["danceName"].ToUpper());
            body.AppendChild(paragraph2);
            addImage("Inline", wordDocument, myTexts["callerPictureFile"], 250, 6.0, 10.0, false);

            
            String[] atoms = myTexts["callerName"].Split('_');
            String country = atoms[2].ToUpper();
            body.AppendChild(GenerateCallerNameParagraph(myTexts["callerName"], country));

            Wp.Paragraph paragraphDanceDates = MyOpenXml.GenerateParagraph(new string[] { danceDates }, new int[] { 24 }, new string[] { "Black" }, false);
            body.AppendChild(paragraphDanceDates);
            body.AppendChild(GenerateDanceLocationParagraph(danceLocation));


            Wp.Paragraph danceFeeParagraph = GenerateFeesParagraph(schemaInfo, danceDateStart);
            body.Append(danceFeeParagraph);

            Wp.Table table = CreateDanceSchemaTable(lang, schemaInfo, schemaName, 14);
            Wp.Paragraph tableParagraph = generateTableParagraph(table, false);

            //tableParagraph.ParagraphProperties.PageBreakBefore = new Wp.PageBreakBefore();
            body.AppendChild(tableParagraph);


            body.AppendChild(GenerateCoffeeParagraph(coffee, 500));
            body.AppendChild(GenerateRotationParagraph());

            var sections = mainPart.Document.Descendants<Wp.SectionProperties>();
            foreach (Wp.SectionProperties sectPr in sections)
            {
                Wp.PageMargin myPageMargin = sectPr.Descendants<Wp.PageMargin>().FirstOrDefault();
                //myPageMargin.Left = myPageMargin.Left / 2;
            }


            String htmlText = htmlStringBuilder.ToString();
            File.WriteAllText(path.Replace("docx", "htm"), htmlText);

            createDemoHtmlFile(path, htmlText);
        }
        private static void createWeekendFlyer(string lang, SchemaInfo schemaInfo, DateTime danceDateStart, string schemaName, string path, bool coffee, string danceLocation, string danceDates, WordprocessingDocument wordDocument)
        {
            MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();

            mainPart.Document = new Wp.Document();
            Wp.Body body = mainPart.Document.AppendChild(new Wp.Body());

            String logoFileName = myTexts["logo_file_name"];
            addImage("Anchor", wordDocument, logoFileName, 84, 10.0, 16.0);
            addImage("Anchor", wordDocument, logoFileName, 84, 178.0, 16.0);
            Wp.Paragraph paragraph1 = GenerateWelcomeParagraphWeekend(myTexts["danceName"].ToUpper(), danceDates);
            body.AppendChild(paragraph1);
            // double scale = 0.7;
            addImage("Inline", wordDocument, myTexts["callerPictureFile"], 250, 6.0, 10.0);

            body.AppendChild(GenerateCallerNameParagraph(myTexts["callerName"]));

            Wp.Table table = CreateDanceSchemaTable(lang, schemaInfo, schemaName, 14);
            Wp.Paragraph tableParagraph = generateTableParagraph(table);
            //body.AppendChild(table);
            body.AppendChild(tableParagraph);


            //body.AppendChild(new Wp.Break());
            body.AppendChild(GenerateDanceLocationParagraph(danceLocation));
            body.AppendChild(GenerateFeesParagraph(schemaInfo, danceDateStart));
            body.AppendChild(GenerateCoffeeParagraph(coffee, 500));
            body.AppendChild(GenerateRotationParagraph());
            MyOpenXml.SetMarginsAndFooter(wordDocument, myTexts["footer"], 10);

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

        private static String createDanceDates(DateTime danceDateStart, DateTime danceDateEnd, out string monthName1, out string monthName2)
        {
            int month1 = danceDateStart.Month;
            int month2 = danceDateEnd.Month;
            DayOfWeek dayOfWeek1 = danceDateStart.DayOfWeek;
            String day1 = DateTimeFormatInfo.CurrentInfo.GetDayName(dayOfWeek1);
            DayOfWeek dayOfWeek2 = danceDateEnd.DayOfWeek;
            String day2 = DateTimeFormatInfo.CurrentInfo.GetDayName(dayOfWeek2);
            monthName1 = DateTimeFormatInfo.CurrentInfo.GetMonthName(month1);
            monthName2 = DateTimeFormatInfo.CurrentInfo.GetMonthName(month2);
            String danceDates = String.Format("{0}-{1} {2} {3}", danceDateStart.Day, danceDateEnd.Day, monthName1, danceDateStart.Year);
            if (month1 != month2)
            {
                danceDates = String.Format("{0}{1}-{2} {3} {4}", danceDateStart.Day, danceDateEnd.Day, monthName1, monthName2, danceDateStart.Year);
            }
            return danceDates;
        }

        private static Wp.Paragraph generateTableLineParagraph()
        {
            Wp.Paragraph paragraph = new Wp.Paragraph();
            Wp.ParagraphProperties paragraphProperties = new Wp.ParagraphProperties();
            Wp.SpacingBetweenLines spacingBetweenLines = new Wp.SpacingBetweenLines();
            spacingBetweenLines.LineRule = Wp.LineSpacingRuleValues.Exact;
            spacingBetweenLines.Line = "350";
            paragraphProperties.Append(spacingBetweenLines);

            Wp.Justification justification = new Wp.Justification() { Val = Wp.JustificationValues.Center };
            paragraphProperties.Append(justification);


            paragraph.Append(paragraphProperties);
            return paragraph;
        }

        private static Wp.Paragraph generateTableParagraph(Wp.Table table, Boolean pageBreakBefore= false)
        {
            Wp.Paragraph paragraph = new Wp.Paragraph();
            Wp.ParagraphProperties paragraphProperties = new Wp.ParagraphProperties();
            Wp.SpacingBetweenLines spacingBetweenLines = new Wp.SpacingBetweenLines();
            spacingBetweenLines.LineRule = Wp.LineSpacingRuleValues.Exact;
            spacingBetweenLines.Line = "200";
            if (pageBreakBefore)
            {
                paragraphProperties.PageBreakBefore = new Wp.PageBreakBefore();
            }

            paragraph.Append(paragraphProperties);
            Wp.Run run = new Wp.Run();
            run.Append(table);
            paragraph.Append(run);
            return paragraph;
        }
        public static Wp.Paragraph GenerateWelcomeParagraphFestival1()
        {
            String[] lines = { myTexts["welcome"].ToUpper() + " " + myTexts["to"].ToUpper() };
            String[] colors = { "Black" };
            int[] fontSizes = { 20 };
            return MyOpenXml.GenerateParagraph(lines, fontSizes, colors, false);
        }
        public static Wp.Paragraph GenerateWelcomeParagraphFestival2(String festivalName)
        {
            String[] lines = { festivalName, "", "" };
            String[] colors = { "Black", "Black", "Black" };
            int[] fontSizes = { 30, 10, 10 };
            return MyOpenXml.GenerateParagraph(lines, fontSizes, colors, false);
        }
        public static Wp.Paragraph GenerateWelcomeParagraphWeekend(String danceName, String danceDates)
        {
            Wp.Paragraph paragraph1 = new Wp.Paragraph();

            String[] lines = { myTexts["welcome"], myTexts["to"], danceName, danceDates };

            String[] colors = { "Black", "Black", "Black", "Black" };
            int[] fontSizes = { 20, 12, 32, 20 };
            return MyOpenXml.GenerateParagraph(lines, fontSizes, colors, false);
        }

        public static Wp.Paragraph GenerateCallerNameParagraph(String callerName, String country = null)
        {
            String[] names = callerName.Split('_');

            String firstName = char.ToUpper(names[0][0]) + names[0].Substring(1);
            String lastName = char.ToUpper(names[1][0]) + names[1].Substring(1);

            callerName = String.Format("{0} {1}", firstName, lastName);
            if (country != null)
            {
            }

            Wp.Table table1 = new Wp.Table();
            MyOpenXml.CreateTableMargins(table1, 150, 50, 0, 50);
            //MyOpenXml.CreateTableBorders(table1, 20, Wp.BorderValues.Double);
            Wp.TableRow tableRow1 = createRow(new string[] { callerName }, new int[] { 5000 }, false, 20);
            table1.Append(tableRow1);
            Wp.TableRow tableRow2 = createRow(new string[] { country }, new int[] { 5000 }, false, 20);
            table1.Append(tableRow2);
            Wp.Paragraph tableParagraph1 = generateTableParagraph(table1);
            return tableParagraph1;

            //return MyOpenXml.GenerateParagraph(lines.ToArray(), fontSizes.ToArray(), colors.ToArray(), borders);
        }


        public static Wp.Paragraph GenerateCallerNameParagraphOld(String callerName, String country = null, Wp.ParagraphBorders borders = null)
        {
            String[] names = callerName.Split('_');

            String firstName = char.ToUpper(names[0][0]) + names[0].Substring(1);
            String lastName = char.ToUpper(names[1][0]) + names[1].Substring(1);

            callerName = String.Format("{0} {1}", firstName, lastName);
            List<String> lines = new List<String>();
            List<String> colors = new List<String>();
            List<int> fontSizes = new List<int>();
            lines.Add(callerName);
            colors.Add("Black");
            fontSizes.Add(32);
            if (country != null)
            {
                lines.Add(country);
                colors.Add("Black");
                fontSizes.Add(32);
            }
            return MyOpenXml.GenerateParagraph(lines.ToArray(), fontSizes.ToArray(), colors.ToArray(), false, borders);
        }

        public static Wp.Paragraph GenerateFeesParagraph(SchemaInfo schemaInfo, DateTime danceDateStart)
        {

            if (schemaInfo.schemaName.StartsWith("weekend"))
            {
                return GenerateWeekendFeesLines(schemaInfo);
            }
            else
            {
                return GenerateFestivalFeesLines(schemaInfo, danceDateStart);
            }

        }
        private static Wp.Paragraph GenerateFestivalFeesLines(SchemaInfo schemaInfo, DateTime danceDateStart)
        {
            List<String> lines = new List<String>();
            lines.Add("Entré");
            String line1 = String.Format(myTexts["festival_fees"], myFees.festival[0], myFees.festival[1],
            myFees.festival[2], myFees.festival[3],
            myFees.festival[4], myFees.festival[5]);
            lines.Add(line1);
            lines.Add(myTexts["festival_fees_text_1"]);
            lines.Add(myTexts["festival_fees_text_2"]);
            lines.Add(myTexts["festival_fees_text_3"]);
            DateTime dueDate1= danceDateStart - new TimeSpan(4 * 24, 0, 0);
            lines.Add(String.Format(myTexts["festival_fees_text_4"], dueDate1.Day +"/"+dueDate1.Month));
            
            lines.Add(myTexts["festival_fees_text_5"]);
            lines.Add(myTexts["festival_fees_text_6"]);

            DateTime dueDate2 = danceDateStart - new TimeSpan(1 * 24, 0, 0);
            lines.Add(String.Format(myTexts["festival_fees_text_7"], dueDate2.Day + "/" + dueDate2.Month));

            lines.Add(myTexts["festival_fees_text_8"]);

            lines.Add(myTexts["festival_fees_text_9"]);
            int[] fontSizes = Enumerable.Repeat(18, 11).ToArray();
            fontSizes[0] = 24;
            String[] colors = Enumerable.Repeat("Black", 11).ToArray();
            return MyOpenXml.GenerateParagraph(lines.ToArray(), fontSizes, colors, true);

        }
        private static Wp.Paragraph GenerateWeekendFeesLines(SchemaInfo schemaInfo)
        {

            List<String> lines = new List<String>();
            if (schemaInfo.schemaName.StartsWith("weekend"))
            {
                String line1 = String.Format(myTexts["weekend_member_fees"], myFees.weekends[0], myFees.weekends[1]);
                htmlStringBuilder.Append("<p class='m8_schema'>\n");
                lines.Add(line1);
                htmlStringBuilder.Append(line1 + "<br>");
                String line2 = String.Format(myTexts["weekend_non_member_fees"], myFees.weekends[2], myFees.weekends[3]);
                lines.Add(line2);
                htmlStringBuilder.Append(line2 + "<br>");
                if (schemaInfo.schemaName == "weekend_january")
                {
                    lines.Add(myTexts["one_pass_sunday"]);
                    htmlStringBuilder.Append(myTexts["one_pass_sunday"] + "<br>");
                }
                String pgPay = myTexts["pg_pay"];
                if (pgPay != "N/A")
                {
                    lines.Add(pgPay);
                    htmlStringBuilder.Append(pgPay + "<br>");
                }
                String swishpay = myTexts["swish_pay"];
                if (swishpay != "N/A")
                {
                    lines.Add(swishpay);
                    htmlStringBuilder.Append(swishpay + "<br>");
                }
            }
            htmlStringBuilder.Append("</p>");

            int[] fontSizes = { 12, 12, 12, 12 };
            String[] colors = { "Black", "Black", "Black", "Black" };
            return MyOpenXml.GenerateParagraph(lines.ToArray(), fontSizes, colors, false);

        }

        public static Wp.Paragraph GenerateDanceLocationParagraph(String danceLocation)
        {
            Wp.Table table1 = new Wp.Table();
            MyOpenXml.CreateTableMargins(table1, 150, 50, 0, 50);
            MyOpenXml.CreateTableBorders(table1, 20, Wp.BorderValues.Double);
            Wp.TableRow tableRow1 = createRow(new string[] { danceLocation }, new int[] { 10000 }, false, 20);
            table1.Append(tableRow1);
            Wp.Paragraph tableParagraph1 = generateTableParagraph(table1);
            return tableParagraph1;
        }
        public static Wp.Paragraph GenerateCoffeeParagraph(Boolean coffee, int maxWidth)
        {
            String text = String.Format("{0} - {1}", myTexts["no_coffee"], myTexts["lunch"]);
            if (coffee)
            {
                text = String.Format("{0} - {1}", myTexts["coffee"], myTexts["lunch"]);
            }
            String[] lines = { text };
            int[] fontSizes = { 13 };
            String[] colors = { "Black" };
            Wp.ParagraphBorders borders = MyOpenXml.createParagraphBorders(Wp.BorderValues.Double, 12);
            htmlStringBuilder.Append(String.Format("<p  class='m8_schema m8_border' style='max-width: {0}px;'>{1}</p>",
                maxWidth, text));

            return MyOpenXml.GenerateParagraph(lines, fontSizes, colors, false, borders);
        }

        public static Wp.Paragraph GenerateRotationParagraph()
        {
            String[] lines = { myTexts["rotation"] };
            int[] fontSizes = { 14 };
            String[] colors = { "Black" };
            htmlStringBuilder.Append("<p  class='m8_schema'>\n");
            htmlStringBuilder.Append(lines[0] + "</p>\n");
            return MyOpenXml.GenerateParagraph(lines, fontSizes, colors, false);
        }

        // Insert a table into a word processing document.
        public static Wp.Table CreateDanceSchemaTable(String lang, SchemaInfo schemaInfo, String schemaName, int fontSize)
        {
            List<DancePass> dancePasses = schemaInfo.danceSchema;

            var n = dancePasses.Select(o => new { Day = o.day }).Distinct();
            int numberOfDistinctDays = n.Count();

            dancePassesDayList.Clear();

            for (int i = 1; i <= numberOfDistinctDays; i++)
            {
                dancePassesDayList.Add(getDancePassesForDay(dancePasses, i));
            }

            if (numberOfDistinctDays == 2)
            {
                htmlStringBuilder.Append("<table class='m8_schema'>");
                Wp.Table table = createWeekendDanceSchemaTable(lang, dancePassesDayList, schemaInfo);
                htmlStringBuilder.Append("</table><br>");
                return table;
            }
            else if (numberOfDistinctDays == 4)
            {
                htmlStringBuilder.Append("<table class='m8_schema'>");
                Wp.Table table = createFestivalDanceSchemaTable(lang, dancePassesDayList, schemaInfo, fontSize);
                htmlStringBuilder.Append("</table><br>");
                return table;
            }
            else
            {
                return null;
            }
        }
        public static Wp.Table createFestivalDanceSchemaTable(String lang, List<DancePass[]> dancePassesDayList, SchemaInfo schemaInfo, int fontSize)
        {
            Wp.Table table = new Wp.Table();
            MyOpenXml.CreateTableBorders(table, 6, Wp.BorderValues.Single, Wp.BorderValues.Single);
            MyOpenXml.CreateTableMargins(table);
            int dayNumber = 0;
            int passNumber = 1;
            foreach (DancePass[] dancePasses in dancePassesDayList)
            {
                foreach (DancePass dancePass in dancePasses)
                {
                    createFestivalRowForFlyer(dayNumber, passNumber, dancePass, table, schemaInfo.colWidth, fontSize);
                    // htmlStringBuilder.Append(createFestivalRowHtml(lang, dayNumber, passNumber, dancePass, schemaInfo));
                    passNumber++;
                }
                dayNumber++;
            }
            IEnumerable<Wp.TableRow> rows = table.Elements<Wp.TableRow>();
            int rowNumber = 0;
            foreach (Wp.TableRow row in rows)
            {
                int cellNumber = 0;
                foreach (Wp.TableCell cell in row.Descendants<Wp.TableCell>())
                {

                    if ((rowNumber == 1 || rowNumber == 4) && cellNumber == 1)
                    {
                        mergeVertical(cell, Wp.MergedCellValues.Restart);
                    }
                    else if ((rowNumber == 2 || rowNumber == 3 || rowNumber == 5) && cellNumber == 1)
                    {
                        mergeVertical(cell, Wp.MergedCellValues.Continue);
                    }
                    cellNumber++;
                }
                rowNumber++;
            }

            return table;

        }
        public static Wp.Table createWeekendDanceSchemaTable(String lang, List<DancePass[]> dancePassesDayList, SchemaInfo schemaInfo)
        {
            Wp.Table table = new Wp.Table();

            createFirstWeekendRow(table, schemaInfo.colWidth.ToArray(), 14);                                 // row 1, Header row
            htmlStringBuilder.Append("<tr class='m8_schema'>" +
               "<th colspan=2 class='m8_schema'>" + myTexts["Saturday"] + "</th>" +
               "<th class='m8_schema m8_space' style='min-width:50px;'></th>" +
                "<th colspan=2 class='m8_schema'>" + myTexts["Sunday"] + "</th>" +
                "</tr>\n");

            createWeekEndRowForFlyer(dancePassesDayList, table, schemaInfo.colWidth, 2);    // row 2
            htmlStringBuilder.Append(createWeekEndRowHtml(lang, dancePassesDayList, 2, schemaInfo));

            // Merge column 1 and 2 if schemaName== "weekend_meeting"
            createWeekEndRowForFlyer(dancePassesDayList, table, schemaInfo.colWidth, 3, schemaInfo.schemaName == "weekend_meeting");   // row 3
            htmlStringBuilder.Append(createWeekEndRowHtml(lang, dancePassesDayList, 3, schemaInfo, schemaInfo.schemaName == "weekend_meeting"));

            if (dancePassesDayList[0].Length > 2 || dancePassesDayList[1].Length > 2)
            {
                createWeekEndRowForFlyer(dancePassesDayList, table, schemaInfo.colWidth, 4);  // row 4
                htmlStringBuilder.Append(createWeekEndRowHtml(lang, dancePassesDayList, 4, schemaInfo));
            }

            return table;

        }
        private static void createFestivalRowForFlyer(int dayNumber, int passNumber, DancePass dancePass, Wp.Table table, List<int> colWidth, int fontSize)
        {
            string weekDay, timeString, level;
            createFestivalRow(dancePass, dayNumber, passNumber,  out weekDay, out timeString, out level);

            String[] content = { passNumber.ToString(), weekDay, timeString, level };

            table.Append(createRow(content, colWidth.ToArray(), false, fontSize));
        }

        private static void createWeekEndRowForFlyer(List<DancePass[]> dancePassesDayList, Wp.Table table, List<int> colWidth, int row, Boolean merge = false)
        {
            string level2, timeString2, level1, timeString1;
            createWeekEndRow(dancePassesDayList, row, out level2, out timeString2, out level1, out timeString1);

            String[] content = { timeString1,
                level1,
                "",
                timeString2,
                level2 };

            table.Append(createRow(content, colWidth.ToArray(), merge));
        }
        private static void createFestivalRow(DancePass dancePass, int dayNumber, int passNumber,  out string weekDay, out string timeString, out string level)
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

        private static void createWeekEndRow(List<DancePass[]> dancePassesDayList, int rowNumber, out string level2, out string timeString2, out string level1, out string timeString1)
        {
            int i1 = rowNumber - 2;
            level2 = "";
            timeString2 = "";
            try
            {
                timeString2 = formatTimeInterval(dancePassesDayList[1][i1]);
                level2 = dancePassesDayList[1][i1].level;
            }
            catch (Exception)
            {

            }
            level1 = "";
            timeString1 = "";
            try
            {
                timeString1 = formatTimeInterval(dancePassesDayList[0][i1]);
                level1 = dancePassesDayList[0][i1].level;
            }
            catch (Exception)
            {

            }
        }

        private static String createWeekEndRowHtml(String lang, List<DancePass[]> dancePassesDayList, int rowNumber, SchemaInfo schemaInfo, Boolean merge = false)
        {
            string level2, timeString2, level1, timeString1;
            createWeekEndRow(dancePassesDayList, rowNumber, out level2, out timeString2, out level1, out timeString1);


            String row = String.Format("<tr class='m8_schema'><td class='m8_schema m8_time'>{0}</td><td class='m8_schema m8_level'>{1}</td><td class='m8_schema m8_space'> </td><td class='m8_schema m8_time'>{2}</td><td class='m8_schema m8_level'>{3}</td></tr>",
                timeString1, level1, timeString2, level2);
            return row;

        }



        /// <summary>
        /// Create the header row for a Weekend dance schedule
        /// </summary>
        /// <param name="table"></param>
        /// <param name="colWidth"></param>
        private static void createFirstWeekendRow(Wp.Table table, int[] colWidth, int fontSize)
        {
            String[] content = new String[5];// = { "Lördag", "", "Söndag", "" };
            content[0] = myTexts["Saturday"];
            content[3] = myTexts["Sunday"];

            Wp.TableRow tr1 = createRow1(content, colWidth, fontSize);
            table.Append(tr1);
        }

        //private static String formatTimeInterval(DancePass[] dancePasses, int row)
        //{
        //    //String text = String.Format("{0}-{1}", dancePasses[row].start_time, dancePasses[row].end_time);
        //    //if (dancePasses[row].end_time.Length > 6) // quick and dirty solution for Årsmöte
        //    //{
        //    //    text = String.Format("{0} {1}", dancePasses[row].start_time, dancePasses[row].end_time);
        //    //}
        //    return formatTimeInterval(dancePasses[row]);
        //}
        private static String formatTimeInterval(DancePass dancePass)
        {
            String text = String.Format("{0}-{1}", dancePass.start_time, dancePass.end_time);
            if (dancePass.end_time.Length > 6) // quick and dirty solution for Årsmöte
            {
                text = String.Format("{0} {1}", dancePass.start_time, dancePass.end_time);
            }
            return text;
        }

        private static DancePass[] getDancePassesForDay(List<DancePass> dancePasses, int day)
        {
            IEnumerable<DancePass> dayOnePasses = from dancePass in dancePasses
                                                  where dancePass.day == day
                                                  orderby dancePass.pass_no ascending
                                                  select dancePass;
            var dayOnePassesArray = dayOnePasses.ToArray();
            return dayOnePassesArray;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="content"></param>
        /// <param name="colWidth">dxa = point/20</param>
        /// <param name="fontSize"></param>
        /// <returns></returns>
        private static Wp.TableRow createRow1(String[] content, int[] colWidth, int fontSize)
        {
            Wp.TableRow row = new Wp.TableRow();
            var rowProps = new Wp.TableRowProperties();


            rowProps.Append(new Wp.TableJustification { Val = Wp.TableRowAlignmentValues.Center });

            row.Append(rowProps);
            Wp.TableCell[] tableCell = new Wp.TableCell[5];
            for (int i = 0; i < content.Length; i++)
            {
                tableCell[i] = createACell(content[i], colWidth[i], true, fontSize);
                row.Append(tableCell[i]);
            }


            MergeCells(tableCell[0], tableCell[1]);
            MergeCells(tableCell[3], tableCell[4]);

            return row;

        }
        private static void MergeCells(Wp.TableCell tc1, Wp.TableCell tc2)
        {
            Wp.TableCellProperties cellOneProperties = new Wp.TableCellProperties();
            cellOneProperties.Append(new Wp.HorizontalMerge()
            {
                Val = Wp.MergedCellValues.Restart
            });

            Wp.TableCellProperties cellTwoProperties = new Wp.TableCellProperties();
            cellTwoProperties.Append(new Wp.HorizontalMerge()
            {
                Val = Wp.MergedCellValues.Continue
            });

            tc1.Append(cellOneProperties);
            tc2.Append(cellTwoProperties);

        }

        /// <summary>
        /// Create a row in the dance schema
        /// </summary>
        /// <param name="content">Array with texts for each cell(column)</param>
        /// <param name="colWidth">Array with columns widths in dxa= point/20</param>
        /// <param name="merge">If true, merges the first two cells in the row</param>
        /// <returns></returns>
        private static Wp.TableRow createRow(String[] content, int[] colWidth, bool merge = false, int fontSize=28)
        {

            Wp.TableRow row = new Wp.TableRow();
            var rowProps = new Wp.TableRowProperties();
            rowProps.Append(new Wp.TableJustification { Val = Wp.TableRowAlignmentValues.Center });

            row.Append(rowProps);
            List<Wp.TableCell> tableCells = new List<Wp.TableCell>();
            for (int i = 0; i < content.Length; i++)
            {
                Wp.TableCell tableCell = createACell(content[i], colWidth[i], false, fontSize);
                tableCells.Add(tableCell);
                row.Append(tableCell);
            }
            if (merge)
            {
                MergeCells(tableCells[0], tableCells[1]);
            }
            return row;

        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="tableCell"></param>
        /// <param name="mergeType">Wp.MergedCellValues.Restart or Wp.MergedCellValues.Continue</param>
        private static void mergeVertical(Wp.TableCell tableCell, Wp.MergedCellValues mergeType)
        {
            var cellProps = new Wp.TableCellProperties();
            cellProps.Append(new Wp.VerticalMerge()
            {
                Val = mergeType
            });

            Wp.TableCellVerticalAlignment tcVA = new Wp.TableCellVerticalAlignment() { Val = Wp.TableVerticalAlignmentValues.Center };
            cellProps.Append(tcVA);

            tableCell.Append(cellProps);

        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="text"></param>
        /// <param name="width">dxa= point/20</param>
        /// <param name="underline"></param>
        /// <param name="fontSize">Font size in points</param>
        /// <returns></returns>
        private static Wp.TableCell createACell(String text, int width, Boolean underline, int fontSize)
        {
            String widthStr = width.ToString();
            Wp.TableCell tableCell = new Wp.TableCell();
            Wp.FontSize wpFontSize = new Wp.FontSize { Val = (2*fontSize).ToString() };  // Size in half points
            Wp.RunFonts runFonts = new Wp.RunFonts { Ascii = "Comic Sans MS" };
            var paragraph = generateTableLineParagraph();

            var run = new Wp.Run();
            var txt = new Wp.Text(text);

            Wp.RunProperties runProperties = new Wp.RunProperties();
            if (underline)
            {
                Wp.Underline ul = new Wp.Underline() { Val = Wp.UnderlineValues.Single };
                runProperties.Append(ul);
            }
            runProperties.Append(wpFontSize);
            runProperties.Append(runFonts);

            run.Append(runProperties);
            run.Append(txt);

            paragraph.Append(run);



            tableCell.Append(paragraph);
            // tableCell.Append(run);

            // Specify the width property of the table cell.
            tableCell.Append(new Wp.TableCellProperties(
                new Wp.TableCellWidth() { Type = Wp.TableWidthUnitValues.Dxa, Width = widthStr }));
           
            // Specify the table cell content.
            //tc1.Append(new Wp.Paragraph(new Wp.Run(new Wp.Text(text))));

            return tableCell;

        }

          public static void addImage(String type, WordprocessingDocument wordprocessingDocument, String fileName, int maxHeight,
              double x0_mm, double y0_mm, Boolean border= false)
        {
            MainDocumentPart mainPart = wordprocessingDocument.MainDocumentPart;
            ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);
            //String fileName = @"D:\Mina dokument\Sqd\Motiv8s\Badge\M8-logo1.gif";
            int iWidth = 0;
            int iHeight = 0;


            // Set a default policy level for the "http:" and "https" schemes.
            HttpRequestCachePolicy policy = new HttpRequestCachePolicy(HttpRequestCacheLevel.Default);
            System.Net.HttpWebRequest.DefaultCachePolicy = policy;
            // Create the request.
            //WebRequest request = WebRequest.Create(uri);
            // Define a cache policy for this request only. 
            HttpRequestCachePolicy noCachePolicy = new HttpRequestCachePolicy(HttpRequestCacheLevel.NoCacheNoStore);

            double scale = 1;

            System.Net.WebRequest request = System.Net.WebRequest.Create(fileName);
            request.CachePolicy = noCachePolicy;
            System.Net.WebResponse response = request.GetResponse();
            using (System.IO.Stream responseStream = response.GetResponseStream())
            {
                // Convert the System.Net.Connectstream responseStream to a System.IO.MemoryStream
                // as imagePart.FeedData (below) will need this to function OK
                using (MemoryStream memoryStream = new MemoryStream())
                {
                    using (Bitmap bmp = new Bitmap(responseStream))
                    {
                        iWidth = bmp.Width;
                        iHeight = bmp.Height;
                        bmp.Save(memoryStream, System.Drawing.Imaging.ImageFormat.Jpeg);
                    }

                    //if (iHeight > maxHeight)
                    //{
                    scale = (double)maxHeight / (double)iHeight;
                    //}
                    memoryStream.Position = 0;
                    imagePart.FeedData(memoryStream);
                }
            }
            if (type == "Anchor")
            {
                AddAnchorImageToBody(wordprocessingDocument, mainPart.GetIdOfPart(imagePart), (int)(scale * iWidth), (int)(scale * iHeight), x0_mm, y0_mm);
            }
            else
            {
                MyOpenXml.AddInlineImageToBody(wordprocessingDocument, mainPart.GetIdOfPart(imagePart), (int)(scale * iWidth), (int)(scale * iHeight), border, fileName);
            }
        }
        public static void AddAnchorImageToBody(WordprocessingDocument wordprocessingDocument, String relationshipId, int iWidth, int iHeight, double x0_mm, double y0_mm)
        {
            var element = MyOpenXml.GetAnchorPicture(relationshipId, x0_mm, y0_mm, iWidth, iHeight);
            // Append the reference to the body. The element should be in a DocumentFormat.OpenXml.Wordprocessing.Run.
            Wp.Paragraph paragraph = new Wp.Paragraph(new Wp.Run(element));

             wordprocessingDocument.MainDocumentPart.Document.Body.AppendChild(paragraph);

        }

    
       }
}
