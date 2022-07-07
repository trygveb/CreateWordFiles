using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Cache;
using OXML = DocumentFormat.OpenXml;
using Wp = DocumentFormat.OpenXml.Wordprocessing;

//using A = DocumentFormat.OpenXml.Drawing;
//using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
//using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace CreateWordFiles
{
    internal class FlyerCreator
    {
        /*
         CSS example

         */
        private static Fees myFees;
        private static Dictionary<String, String> myTexts;
        private static Wp.Color wpColorBlackx = new Wp.Color() { Val = "000000" };
        private static Wp.Color wpColorRedx = new Wp.Color() { Val = "FF0000" };
        //public static StringBuilder htmlStringBuilder = new StringBuilder();

        public static void AddAnchorImageToBody(WordprocessingDocument wordprocessingDocument, String relationshipId, int iWidth, int iHeight, double x0_mm, double y0_mm)
        {
            var element = MyOpenXml.GetAnchorPicture(relationshipId, x0_mm, y0_mm, iWidth, iHeight);
            // Append the reference to the body. The element should be in a DocumentFormat.OpenXml.Wordprocessing.Run.
            Wp.Paragraph paragraph = new Wp.Paragraph(new Wp.Run(element));

            wordprocessingDocument.MainDocumentPart.Document.Body.AppendChild(paragraph);
        }

        public static void AddImage(String type, WordprocessingDocument wordprocessingDocument, String fileName, int maxHeight,
                    double x0_mm, double y0_mm, Boolean border = false)
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

        public static Wp.Table CreateDanceSchemaTable(String lang, SchemaInfo schemaInfo, String schemaName, int fontSize, int lineSpace)
        {
            int numberOfDistinctDays = Utility.createDancePassesDaylist(schemaInfo);

            if (numberOfDistinctDays == 2)
            {
                //htmlStringBuilder.Append("<table class='m8_schema'>");
                Wp.Table table = CreateWeekendDanceSchemaTable(lang, Utility.dancePassesDayList, schemaInfo, fontSize, lineSpace);
                //htmlStringBuilder.Append("</table><br>");
                return table;
            }
            else if (numberOfDistinctDays == 4)
            {
                //htmlStringBuilder.Append("<table class='m8_festival'>");
                Wp.Table table = CreateFestivalDanceFlyerSchemaTable(lang, Utility.dancePassesDayList, schemaInfo, fontSize, lineSpace);
                //htmlStringBuilder.Append("</table><br>");
                return table;
            }
            else
            {
                return null;
            }
        }

        public static Wp.Table CreateFestivalDanceFlyerSchemaTable(String lang, List<DancePass[]> dancePassesDayList, SchemaInfo schemaInfo,
                       int fontSize, int lineSpace)
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
                    createFestivalRowForFlyer(dayNumber, passNumber, dancePass, table, schemaInfo.colWidth, fontSize, lineSpace);
                    //htmlStringBuilder.Append(createFestivalRowHtml(lang, dayNumber, passNumber, dancePass, schemaInfo));
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

        public static Wp.Table CreateWeekendDanceSchemaTable(String lang, List<DancePass[]> dancePassesDayList, SchemaInfo schemaInfo, int fontSize, int lineSpace)
        {
            Wp.Table table = new Wp.Table();

            createFirstWeekendRow(table, schemaInfo.colWidth.ToArray(), fontSize, lineSpace);                                 // row 1, Header row

            createWeekEndRowForFlyer(dancePassesDayList, table, schemaInfo.colWidth, 2, false, fontSize, lineSpace);    // row 2

            // Merge column 1 and 2 if schemaName== "weekend_meeting"
            createWeekEndRowForFlyer(dancePassesDayList, table, schemaInfo.colWidth, 3, schemaInfo.schemaName == "weekend_meeting",
                fontSize, lineSpace);   // row 3

            if (dancePassesDayList[0].Length > 2 || dancePassesDayList[1].Length > 2)
            {
                createWeekEndRowForFlyer(dancePassesDayList, table, schemaInfo.colWidth, 4, false, fontSize, lineSpace);  // row 4
            }

            return table;
        }

        /// <summary>
        /// Only weekend dances supported
        /// </summary>
        /// <param name="filepath"></param>
        /// <param name="danceName"></param>
        /// <param name="danceDateStart"></param>
        /// <param name="danceDateEnd"></param>
        public static void CreateWordprocessingDocument(Dictionary<String, String> texts, String lang, SchemaInfo schemaInfo,
            String schemaName, String path, Fees fees, Boolean coffee, String danceLocation, DateTime danceDateStart,
            DateTime danceDateEnd, String extra, List<String> festivalFeesText)
        {
            string monthName1, monthName2, danceDates;

            myTexts = texts;
            myFees = fees;
 

            danceDates = createDanceDates(danceDateStart, danceDateEnd, lang, out monthName1, out monthName2);

            //htmlStringBuilder.Clear();

            try
            {
                using (WordprocessingDocument wordDocument =
                    WordprocessingDocument.Create(path, OXML.WordprocessingDocumentType.Document))
                {
                    if (schemaName.Contains("festival"))
                    {
                        createFestivalFlyer(lang, schemaInfo, danceDateStart, schemaName, path, coffee, danceLocation, danceDates, wordDocument, extra, festivalFeesText);
                    }
                    else
                    {
                        createWeekendFlyer(lang, schemaInfo, danceDateStart, schemaName, path, coffee, danceLocation, danceDates, wordDocument, festivalFeesText);
                    }
                }
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public static Wp.Paragraph GenerateBlankLines(int numberOfLines, int before, int after, int fontSize = 14)
        {
            String[] lines = Enumerable.Repeat(" ", numberOfLines).ToArray();

            int[] fontSizes = Enumerable.Repeat(fontSize, numberOfLines).ToArray();
            String[] colors = Enumerable.Repeat("Black", numberOfLines).ToArray();

            return MyOpenXml.GenerateParagraph(lines, fontSizes, colors, false, null, before, after, false);
        }

        public static Wp.Paragraph GenerateCallerNameParagraph(String callerName, Boolean addCountry = false, int fontSize = 16)
        {
            String[] names = callerName.Split('_');

            String firstName = char.ToUpper(names[0][0]) + names[0].Substring(1);
            String lastName = char.ToUpper(names[1][0]) + names[1].Substring(1);
            callerName = String.Format("{0} {1}", firstName, lastName);

            Wp.Table table1 = new Wp.Table();
            //htmlStringBuilder.Append("<table class='m8_schema'>\n");
            MyOpenXml.CreateTableMargins(table1, 150, 50, 0, 50);
            Wp.TableRow tableRow1 = createRow(new string[] { callerName }, new int[] { 5000 }, false, fontSize, 400);
            table1.Append(tableRow1);
            if (addCountry)
            {
                String country = char.ToUpper(names[2][0]) + names[2].Substring(1);
                Wp.TableRow tableRow2 = createRow(new string[] { country }, new int[] { 5000 }, false, fontSize, 400);
                table1.Append(tableRow2);
            }
            Wp.Paragraph tableParagraph1 = generateTableParagraph(table1, false, 0, 0);
            //htmlStringBuilder.Append("</table'>\n");
            //String test = htmlStringBuilder.ToString();
            return tableParagraph1;

            //return MyOpenXml.GenerateParagraph(lines.ToArray(), fontSizes.ToArray(), colors.ToArray(), borders);
        }

        public static Wp.Paragraph GenerateCoffeeParagraph(Boolean coffee, int margin, Boolean merge, int columnWidth,
            int fontSize, int lineSpace, int distanceBefore, int distanceAfter)
        {
            Wp.Table table1 = new Wp.Table();
            //htmlStringBuilder.Append("<table class='m8_schema'>\n");
            String text1 = myTexts["no_coffee"];
            if (coffee)
            {
                text1 = myTexts["coffee"];
            }

            //Start- and end margin always the same in order to center the text
            MyOpenXml.CreateTableMargins(table1, 150, margin, 0, margin);
            MyOpenXml.CreateTableBorders(table1, 20, Wp.BorderValues.Double);
            if (merge)
            {
                Wp.TableRow tableRow1 = createRow(new string[] { text1 + "  " + myTexts["lunch"] }, new int[] { columnWidth }, false, fontSize, lineSpace);
                table1.Append(tableRow1);
            }
            else
            {
                Wp.TableRow tableRow1 = createRow(new string[] { text1 }, new int[] { columnWidth }, false, fontSize, lineSpace);
                Wp.TableRow tableRow2 = createRow(new string[] { myTexts["lunch"] }, new int[] { columnWidth }, false, fontSize, lineSpace);
                table1.Append(tableRow1);
                table1.Append(tableRow2);
            }
            //htmlStringBuilder.Append("</table'>\n");
            Wp.Paragraph tableParagraph1 = generateTableParagraph(table1, false, distanceBefore, distanceAfter);
            return tableParagraph1;
        }

        public static Wp.Paragraph GenerateDanceLocationParagraph(String danceLocation, int fontSize = 20, int lineSpace = 400, int columnWidth = 7500)
        {
            Wp.Table table1 = new Wp.Table();
            MyOpenXml.CreateTableMargins(table1, 150, 50, 50, 50);
            MyOpenXml.CreateTableBorders(table1, 20, Wp.BorderValues.Double);
            Wp.TableRow tableRow1 = createRow(new string[] { danceLocation }, new int[] { columnWidth }, false, fontSize, lineSpace);
            table1.Append(tableRow1);
            Wp.Paragraph tableParagraph1 = generateTableParagraph(table1, false, 0, 00);
            //htmlStringBuilder.Append(String.Format("<p class='m8_schema m8_border' style='max-width: 450px;'>{0}</p>", danceLocation));
            return tableParagraph1;
        }

        public static Wp.Paragraph GenerateLastpageParagraph(String extra)
        {
            String[] lines = { myTexts["more_info_1"], myTexts["more_info_2"], " ", myTexts["more_info_3"],
              " ",   myTexts["more_info_4"]," ", " ",  extra , " ", " " , " ", " " };
            int[] fontSizes = Enumerable.Repeat(20, lines.Length).ToArray();
            String[] colors = Enumerable.Repeat("Black", lines.Length).ToArray();
            return MyOpenXml.GenerateParagraph(lines, fontSizes, colors, true, null);
        }

        public static Wp.Paragraph GenerateRotationParagraph(int before, int after, int fontSize = 14)
        {
            String[] lines = { myTexts["rotation"] };
            int[] fontSizes = { fontSize };
            String[] colors = { "Black" };
            //htmlStringBuilder.Append("<p  class='m8_schema'>\n");
            //htmlStringBuilder.Append(lines[0] + "</p>\n");
            return MyOpenXml.GenerateParagraph(lines, fontSizes, colors, false, null, before, after);
        }

        public static Wp.Paragraph GenerateWelcomeParagraphFestival1()
        {
            String[] lines = { myTexts["welcome"].ToUpper() + " " + myTexts["to"].ToUpper() };
            String[] colors = { "Black" };
            int[] fontSizes = { 20 };
            return MyOpenXml.GenerateParagraph(lines, fontSizes, colors, false, null, 0, 160);
        }

        public static Wp.Paragraph GenerateWelcomeParagraphFestival2(String festivalName)
        {
            String[] lines = { festivalName, "", "" };
            String[] colors = { "Black", "Black", "Black" };
            int[] fontSizes = { 30, 10, 10 };
            return MyOpenXml.GenerateParagraph(lines, fontSizes, colors, false, null);
        }

        public static Wp.Paragraph GenerateWelcomeParagraphWeekend(String danceName, String danceDates, int distanceBefore, int distanceAfter)
        {
            String[] lines = { myTexts["welcome"], myTexts["to"], danceName, danceDates };

            String[] colors = { "Black", "Black", "Black", "Black" };
            int[] fontSizes = { 20, 12, 32, 20 };
            return MyOpenXml.GenerateParagraph(lines, fontSizes, colors, false, null, distanceBefore, distanceAfter);
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="text"></param>
        /// <param name="width">dxa= point/20</param>
        /// <param name="underline"></param>
        /// <param name="fontSize">Font size in points</param>
        /// <returns></returns>
        private static Wp.TableCell createACell(String text, int width, Boolean underline, int fontSize,
            int lineSpace)
        {
            String widthStr = width.ToString();
            Wp.TableCell tableCell = new Wp.TableCell();
            Wp.FontSize wpFontSize = new Wp.FontSize { Val = (2 * fontSize).ToString() };  // Size in half points
            Wp.RunFonts runFonts = new Wp.RunFonts { Ascii = "Comic Sans MS" };
            var paragraph = generateTableLineParagraph(lineSpace);

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
            Wp.TableCellProperties x = new Wp.TableCellProperties();
            tableCell.Append(x);
            Wp.TableCellMargin y = new Wp.TableCellMargin();
            x.Append(y);
            y.BottomMargin = new Wp.BottomMargin() { Width = "50" };
            y.TopMargin = new Wp.TopMargin() { Width = "50" };
            x.TableCellWidth = new Wp.TableCellWidth() { Type = Wp.TableWidthUnitValues.Dxa, Width = widthStr };
            // Specify the width property of the table cell.
            //tableCell.Append(new Wp.TableCellProperties(
            //     new Wp.TableCellWidth() { Type = Wp.TableWidthUnitValues.Dxa, Width = widthStr }));

            return tableCell;
        }

        private static String createDanceDates(DateTime danceDateStart, DateTime danceDateEnd, String lang, out string monthName1, out string monthName2) 
        {
            int month1 = danceDateStart.Month;
            int month2 = danceDateEnd.Month;
            DateTimeFormatInfo dateTimeFormatInfo = Utility.dateTimeFormatInfo_se;

            //DateTimeFormatInfo dateTimeFormatInfo = CultureInfo.CurrentCulture.DateTimeFormat;
            if (lang == "en")
            {
                dateTimeFormatInfo = Utility.dateTimeFormatInfo_en;
            }

            //DayOfWeek dayOfWeek1 = danceDateStart.DayOfWeek;
            //String day1 = DateTimeFormatInfo.CurrentInfo.GetDayName(dayOfWeek1);
            //DayOfWeek dayOfWeek2 = danceDateEnd.DayOfWeek;
            //String day2 = DateTimeFormatInfo.CurrentInfo.GetDayName(dayOfWeek2);
            monthName1 = dateTimeFormatInfo.GetMonthName(month1);
            monthName2 = dateTimeFormatInfo.GetMonthName(month2);
            String danceDates = String.Format("{0}-{1} {2} {3}", danceDateStart.Day, danceDateEnd.Day, monthName1, danceDateStart.Year);
            if (lang == "en")
            {
                danceDates = String.Format(" {2} {0}-{1} {3}", danceDateStart.Day, danceDateEnd.Day, monthName1, danceDateStart.Year);
            }
            if (month1 != month2)
            {
                danceDates = String.Format("{0}{1}-{2}{3} {4}", danceDateStart.Day,  monthName1, danceDateEnd.Day,  monthName2, danceDateStart.Year);
                if (lang == "en")
                {
                    danceDates = String.Format("{1} {0}-{3} {2} {4}", danceDateStart.Day, monthName1, danceDateEnd.Day, monthName2, danceDateStart.Year);

                }
            }
             return danceDates;
        }

        private static void createFestivalFlyer(string lang, SchemaInfo schemaInfo, DateTime danceDateStart,
                                                                                                    string schemaName, string path, bool coffee, string danceLocation, string danceDates,
            WordprocessingDocument wordDocument, String extra, List<String> festivalFeesText)
        {
            MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();

            mainPart.Document = new Wp.Document();
            Wp.Body body = mainPart.Document.AppendChild(new Wp.Body());
            //MyOpenXml.SetMarginsAndFooter(wordDocument, myTexts["footer"], 10, 1500, 2000, 2000);

            MyOpenXml.SetMarginsAndFooter(wordDocument,
                String.Format(myTexts["footer"], DateTime.UtcNow.ToString("yyyy-MM-dd")), 11, 1000);

            Wp.Paragraph paragraph1 = GenerateWelcomeParagraphFestival1();
            body.AppendChild(paragraph1);


            AddImage("Inline", wordDocument, myTexts["logo_file_name"], 200, 6.0, 10.0);
            Wp.Paragraph paragraph2 = GenerateWelcomeParagraphFestival2(myTexts["danceName"].ToUpper());
            body.AppendChild(paragraph2);

            AddImage("Inline", wordDocument, myTexts["callerPictureFile"], 259, 6.0, 10.0, false);
            //Wp.Paragraph callerParagraph = AddCallerPicture.GenerateParagraph();
            //body.AppendChild(callerParagraph);

            body.AppendChild(GenerateCallerNameParagraph(myTexts["callerName"], true, 20));

            Wp.Paragraph paragraphDanceDates = generateDanceDatesParagraph(danceDates);
            body.AppendChild(paragraphDanceDates);

            body.AppendChild(GenerateDanceLocationParagraph(danceLocation, 20, 600, 11000));

            Wp.Paragraph paragraphEntre = generateFeeHeadlineParagraph();
            body.AppendChild(paragraphEntre);

           generateFestivalFeesParagraph(schemaInfo, danceDateStart, festivalFeesText, mainPart, body);

            body.AppendChild(GenerateBlankLines(1, 0, 0, 14));

            Wp.Paragraph paragraphProgramTitle = generateProgramTitleParagraph();
            body.AppendChild(paragraphProgramTitle);

            body.AppendChild(GenerateBlankLines(1, 0, 0, 14));

            Wp.Table table = CreateDanceSchemaTable(lang, schemaInfo, schemaName, 24, 800);
            Wp.Paragraph tableParagraph = generateTableParagraph(table, false, 0, 0);

            //tableParagraph.ParagraphProperties.PageBreakBefore = new Wp.PageBreakBefore();
            body.AppendChild(tableParagraph);
            if (schemaInfo.schemaName.Contains("reduced_C3b"))
            {
                body.AppendChild(MyOpenXml.GenerateParagraph(new string[] { myTexts["reduced_C3b"] }, new int[] { 18 },
                    new string[] { "Black" }, false, null, 0, 0));
            }
            body.AppendChild(GenerateRotationParagraph(100, 300, 20));

            body.AppendChild(GenerateBlankLines(2, 0, 0, 14));

            body.AppendChild(GenerateCoffeeParagraph(coffee, 300, false, 10000, 20, 400, 0, 0));
            //htmlStringBuilder.Append("<p class='m8 m8_border'>\n");
            String text1 = myTexts["no_coffee"];
            if (coffee)
            {
                text1 = myTexts["coffee"];
            }

            //htmlStringBuilder.Append(text1 + " - " + myTexts["lunch"]);
            //htmlStringBuilder.Append("</p'>\n");
            body.AppendChild(GenerateLastpageParagraph(extra));

            AddImage("Inline", wordDocument, myTexts["logo_file_name"], 200, 6.0, 10.0);
            AddImage("Anchor", wordDocument, myTexts["saasdc_logo_file_name"], 100, 180.0, 220.0, false);

            var sections = mainPart.Document.Descendants<Wp.SectionProperties>();
            foreach (Wp.SectionProperties sectPr in sections)
            {
                Wp.PageMargin myPageMargin = sectPr.Descendants<Wp.PageMargin>().FirstOrDefault();
                //myPageMargin.Left = myPageMargin.Left / 2;
            }
        }

        private static void createFestivalRowForFlyer(int dayNumber, int passNumber, DancePass dancePass, Wp.Table table,
            List<int> colWidth, int fontSize, int lineSpace)
        {
            string weekDay, timeString, level;
            Utility.createFestivalRow(dancePass, dayNumber, passNumber, out weekDay, out timeString, out level);

            String[] content = { passNumber.ToString(), weekDay, timeString, level };

            table.Append(createRow(content, colWidth.ToArray(), false, fontSize, lineSpace));
        }

        /// <summary>
        /// Create the header row for a Weekend dance schedule
        /// </summary>
        /// <param name="table"></param>
        /// <param name="colWidth"></param>
        private static void createFirstWeekendRow(Wp.Table table, int[] colWidth, int fontSize, int lineSpace)
        {
            String[] content = new String[5];// = { "Lördag", "", "Söndag", "" };
            content[0] = myTexts["Saturday"];
            content[3] = myTexts["Sunday"];

            Wp.TableRow tr1 = createRow1(content, colWidth, fontSize, lineSpace);
            table.Append(tr1);
        }

        /// <summary>
        /// Create a row in the dance schema
        /// </summary>
        /// <param name="content">Array with texts for each cell(column)</param>
        /// <param name="colWidth">Array with columns widths in dxa= point/20</param>
        /// <param name="merge">If true, merges the first two cells in the row</param>
        /// <returns></returns>
        private static Wp.TableRow createRow(String[] content, int[] colWidth, bool merge, int fontSize, int lineSpace)
        {
            Wp.TableRow row = new Wp.TableRow();
            var rowProps = new Wp.TableRowProperties();
            rowProps.Append(new Wp.TableJustification { Val = Wp.TableRowAlignmentValues.Center });

            row.Append(rowProps);
            List<Wp.TableCell> tableCells = new List<Wp.TableCell>();
            //if (createHtml)
            //{
            //    htmlStringBuilder.Append("<tr class='m8_festival'>\n");
            //}
            for (int i = 0; i < content.Length; i++)
            {
                Wp.TableCell tableCell = createACell(content[i], colWidth[i], false, fontSize, lineSpace);
                //if (createHtml)
                //{
                //    FlyerCreator.htmlStringBuilder.Append(String.Format("<td class='m8_festival'>{0}</td>", content[i]));
                //}
                tableCells.Add(tableCell);
                row.Append(tableCell);
            }
            if (merge)
            {
                mergeCells(tableCells[0], tableCells[1]);
            }
            //if (createHtml)
            //{
            //    FlyerCreator.htmlStringBuilder.Append("</tr>\n");
            //}
            return row;
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="content"></param>
        /// <param name="colWidth">dxa = point/20</param>
        /// <param name="fontSize"></param>
        /// <returns></returns>
        private static Wp.TableRow createRow1(String[] content, int[] colWidth, int fontSize, int lineSpace)
        {
            Wp.TableRow row = new Wp.TableRow();
            var rowProps = new Wp.TableRowProperties();

            rowProps.Append(new Wp.TableJustification { Val = Wp.TableRowAlignmentValues.Center });

            row.Append(rowProps);
            Wp.TableCell[] tableCell = new Wp.TableCell[5];
            for (int i = 0; i < content.Length; i++)
            {
                tableCell[i] = createACell(content[i], colWidth[i], true, fontSize, lineSpace);
                row.Append(tableCell[i]);
            }

            mergeCells(tableCell[0], tableCell[1]);
            mergeCells(tableCell[3], tableCell[4]);

            return row;
        }

        private static void createWeekendFlyer(string lang, SchemaInfo schemaInfo, DateTime danceDateStart, string schemaName,
            string path, bool coffee, string danceLocation, string danceDates, WordprocessingDocument wordDocument, List<String> festivalFeesText)
        {
            MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();

            mainPart.Document = new Wp.Document();
            Wp.Body body = mainPart.Document.AppendChild(new Wp.Body());

            String logoFileName = myTexts["logo_file_name"];
            AddImage("Anchor", wordDocument, logoFileName, 95, 10.0, 16.0);
            AddImage("Anchor", wordDocument, logoFileName, 95, 178.0, 16.0);
            Wp.Paragraph paragraph1 = GenerateWelcomeParagraphWeekend(myTexts["danceName"].ToUpper(), danceDates, 0, 480);
            body.AppendChild(paragraph1);
            // double scale = 0.7;
            AddImage("Inline", wordDocument, myTexts["callerPictureFile"], 250, 6.0, 10.0);

            body.AppendChild(GenerateCallerNameParagraph(myTexts["callerName"], false));

            Wp.Table table = CreateDanceSchemaTable(lang, schemaInfo, schemaName, 14, 350);
            //  Wp.Paragraph tableParagraph = generateTableParagraph(table, false);
            //  body.AppendChild(tableParagraph);
            body.AppendChild(table);
            if (schemaInfo.schemaName.Contains("reduced_C3b"))
            {
                body.AppendChild(MyOpenXml.GenerateParagraph(new string[] { myTexts["reduced_C3b"] }, new int[] { 14 },
                    new string[] { "Black" }, false, null, 0, 0));
            }

            body.AppendChild(GenerateRotationParagraph(60, 360));
            body.AppendChild(GenerateDanceLocationParagraph(danceLocation, 14, 390));
            body.AppendChild(generateWeekendFeesLines(schemaInfo));

            body.AppendChild(GenerateCoffeeParagraph(coffee, 50, true, 10000, 14, 320, 0, 0));
            MyOpenXml.SetMarginsAndFooter(wordDocument, myTexts["footer"], 10);

            //String htmlText = htmlStringBuilder.ToString();
            //File.WriteAllText(path.Replace("docx", "htm"), htmlText);

            //createDemoHtmlFile(path, htmlText);
        }

        private static void createWeekEndRowForFlyer(List<DancePass[]> dancePassesDayList, Wp.Table table,
            List<int> colWidth, int row, Boolean merge = false, int fontSize = 14, int lineSpace = 400)
        {
            string level2, timeString2, level1, timeString1;
            Utility.createWeekEndRow(dancePassesDayList, row, out level2, out timeString2, out level1, out timeString1);

            String[] content = { timeString1,
                level1,
                "",
                timeString2,
                level2 };

            table.Append(createRow(content, colWidth.ToArray(), merge, fontSize, lineSpace));
        }

        private static Wp.Paragraph generateDanceDatesParagraph(string danceDates)
        {
            // htmlStringBuilder.Append(String.Format("<p class='m8'>{0}</p>\n", danceDates));
            return MyOpenXml.GenerateParagraph(new string[] { danceDates }, new int[] { 24 },
                new string[] { "Black" }, false, null, 0, 160);
        }

        private static Wp.Paragraph generateFeeHeadlineParagraph()
        {
            //  htmlStringBuilder.Append("<p class='m8'><span class='m8_headline'>Entré</span></p>\n");

            return MyOpenXml.GenerateParagraph(
                new string[] { "Entré" }, new int[] { 24 }, new string[] { "Black" }, true, null, 0, 120, true);
        }

        private static void generateFestivalFeesParagraph(SchemaInfo schemaInfo, DateTime danceDateStart, List<string> festivalFeeTexts, MainDocumentPart mainPart, Wp.Body body)
        {
            List<String> linesWithoutBullets = new List<String>();
            List<String> linesWithBullets = new List<String>();

            // festivalFeeTexts may have zero or one bullet lists
            List<string> festivalFeeTextsWithoutBullets;
            List<string> festivalFeeTextsWithBullets;

            Utility.findBullets(festivalFeeTexts, out festivalFeeTextsWithBullets, out festivalFeeTextsWithoutBullets);
            Utility.generateFestivalFeeLines(danceDateStart, festivalFeeTextsWithoutBullets, myFees, linesWithoutBullets);
            int[] fontSizes = Enumerable.Repeat(18, linesWithoutBullets.Count).ToArray();
            String[] colors = Enumerable.Repeat("Black", linesWithoutBullets.Count).ToArray();
            Wp.Paragraph paragraph = MyOpenXml.GenerateParagraph(linesWithoutBullets.ToArray(), fontSizes, colors, false, null, 0, 0, false);
            body.Append(paragraph);

            if (festivalFeeTextsWithBullets.Count > 0)
            {
                Utility.generateFestivalFeeLines(danceDateStart, festivalFeeTextsWithBullets, myFees, linesWithBullets);
                MyOpenXml.AddBulletList(linesWithBullets, mainPart, body, 18);
            }
        }

        private static Wp.Paragraph generateProgramTitleParagraph()
        {
            return MyOpenXml.GenerateParagraph(new string[]
            { "PROGRAM" }, new int[] { 24 }, new string[] { "Black" }, true, null, 0, 120, true);
        }

        private static Wp.Paragraph generateTableLineParagraph(int lineSpacing = 400)
        {
            Wp.Paragraph paragraph = new Wp.Paragraph();
            Wp.ParagraphProperties paragraphProperties = new Wp.ParagraphProperties();
            Wp.SpacingBetweenLines spacingBetweenLines = new Wp.SpacingBetweenLines();
            spacingBetweenLines.LineRule = Wp.LineSpacingRuleValues.Exact;
            spacingBetweenLines.After = "0";
            spacingBetweenLines.Before = "0";
            spacingBetweenLines.Line = lineSpacing.ToString();
            paragraphProperties.Append(spacingBetweenLines);

            Wp.Justification justification = new Wp.Justification() { Val = Wp.JustificationValues.Center };
            paragraphProperties.Append(justification);

            paragraph.Append(paragraphProperties);
            return paragraph;
        }

        /// <summary>
        /// Used to add PageBreakBefore if neccesary
        /// </summary>
        /// <param name="table"></param>
        /// <param name="pageBreakBefore"></param>
        /// <returns></returns>
        private static Wp.Paragraph generateTableParagraph(Wp.Table table, Boolean pageBreakBefore, int distanceBefore, int distanceAfter)
        {
            Wp.Paragraph paragraph = new Wp.Paragraph();
            Wp.ParagraphProperties paragraphProperties = new Wp.ParagraphProperties();
            if (pageBreakBefore)
            {
                paragraphProperties.PageBreakBefore = new Wp.PageBreakBefore();
            }
            Wp.SpacingBetweenLines spacingBetweenLines = new Wp.SpacingBetweenLines();
            spacingBetweenLines.LineRule = Wp.LineSpacingRuleValues.Exact;
            spacingBetweenLines.Before = distanceBefore.ToString();
            spacingBetweenLines.After = distanceAfter.ToString();
            paragraphProperties.Append(spacingBetweenLines);

            paragraph.Append(paragraphProperties);
            Wp.Run run = new Wp.Run();
            run.Append(table);
            paragraph.Append(run);
            return paragraph;
        }

        private static Wp.Paragraph generateWeekendFeesLines(SchemaInfo schemaInfo)
        {
            List<String> lines = new List<String>();
            if (schemaInfo.schemaName.StartsWith("weekend"))
            {
                String line1 = String.Format(myTexts["weekend_member_fees"], myFees.weekends[0], myFees.weekends[1]);
                lines.Add(line1);
                String line2 = String.Format(myTexts["weekend_non_member_fees"], myFees.weekends[2], myFees.weekends[3]);
                lines.Add(line2);
                if (schemaInfo.schemaName == "weekend_january")
                {
                    lines.Add(myTexts["one_pass_sunday"]);
                }
                String pgPay = myTexts["weekend_pg_pay"];
                if (pgPay != "N/A")
                {
                    lines.Add(pgPay);
                }
                String swishpay = myTexts["weekend_swish_pay"];
                if (swishpay != "N/A")
                {
                    lines.Add(swishpay);
                }
            }

            //int[] fontSizes = { 12, 12, 12, 12 };
            int[] fontSizes = Enumerable.Repeat(12, lines.Count).ToArray();
            String[] colors = Enumerable.Repeat("Black", lines.Count).ToArray();
            return MyOpenXml.GenerateParagraph(lines.ToArray(), fontSizes, colors, false, null, 0, 240);
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
        private static void mergeCells(Wp.TableCell tc1, Wp.TableCell tc2)
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
    }
}