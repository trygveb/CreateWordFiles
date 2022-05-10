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
        private static StringBuilder htmlStringBuilder=new StringBuilder();

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
                    // Add a main document part. 
                    MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();

                    mainPart.Document = new Wp.Document();
                    Wp.Body body = mainPart.Document.AppendChild(new Wp.Body());

                    String logoFileName = "https://motiv8s.se/19/images/M8/Logga_Transparent.jpg";
                    addImage("Anchor", wordDocument, logoFileName, 128, 10.0, 16.0);
                    Wp.Paragraph paragraph1 = GenerateWelcomeParagraph(myTexts["danceName"].ToUpper(), danceDates);
                    body.AppendChild(paragraph1);
                    // double scale = 0.7;
                    addImage("Inline", wordDocument, myTexts["callerPictureFile"], 275, 6.0, 10.0);

                    body.AppendChild(GenerateCallerNameParagraph(myTexts["callerName"]));

                    Wp.Table table = CreateDanceSchemaTable(lang, schemaInfo, schemaName);
                    Wp.Paragraph tableParagraph = generateTableParagraph(table);
                    //body.AppendChild(table);
                    body.AppendChild(tableParagraph);


                    //body.AppendChild(new Wp.Break());
                    body.AppendChild(GenerateDanceLocationParagraph(danceLocation, 400));
                    body.AppendChild(GenerateFeesParagraph(schemaInfo));
                    body.AppendChild(GenerateCoffeeParagraph(coffee, 500));
                    body.AppendChild(GenerateRotationParagraph());
                    ApplyFooter(wordDocument);

                    String htmlText = htmlStringBuilder.ToString();
                    File.WriteAllText(path.Replace("docx", "htm"), htmlText);

                    createDemoHtmlFile(path, htmlText);
                }
            }
            catch (Exception e)
            {
                throw e;
            }
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

        public static void ApplyFooter(WordprocessingDocument doc, String footerText="Footer")
        {
            // Get the main document part.
            MainDocumentPart mainDocPart = doc.MainDocumentPart;

            FooterPart footerPart1 = mainDocPart.AddNewPart<FooterPart>("r98");



            Wp.Footer footer1 = new Wp.Footer();

            Wp.Paragraph paragraph1 = new Wp.Paragraph() { };



            Wp.Run run1 = new Wp.Run();
            Wp.Text text1 = new Wp.Text();
            text1.Text = footerText;

            run1.Append(text1);

            paragraph1.Append(run1);


            footer1.Append(paragraph1);

            footerPart1.Footer = footer1;



            Wp.SectionProperties sectionProperties1 = mainDocPart.Document.Body.Descendants<Wp.SectionProperties>().FirstOrDefault();
            if (sectionProperties1 == null)
            {
                sectionProperties1 = new Wp.SectionProperties() { };
                mainDocPart.Document.Body.Append(sectionProperties1);
            }
            Wp.FooterReference footerReference1 = new Wp.FooterReference() { Type = DocumentFormat.OpenXml.Wordprocessing.HeaderFooterValues.Default, Id = "r98" };


            sectionProperties1.InsertAt(footerReference1, 0);

        }
        private static Wp.Paragraph generateTableLineParagraph()
        {
            Wp.Paragraph paragraph = new Wp.Paragraph();
            Wp.ParagraphProperties paragraphProperties = new Wp.ParagraphProperties();
            Wp.SpacingBetweenLines spacingBetweenLines = new Wp.SpacingBetweenLines();
            spacingBetweenLines.LineRule = Wp.LineSpacingRuleValues.Exact;
            spacingBetweenLines.Line = "270";
            paragraphProperties.Append(spacingBetweenLines);
            paragraph.Append(paragraphProperties);
            return paragraph;
        }

        private static Wp.Paragraph generateTableParagraph(Wp.Table table)
        {
            Wp.Paragraph paragraph = new Wp.Paragraph();
            Wp.ParagraphProperties paragraphProperties = new Wp.ParagraphProperties();
            Wp.SpacingBetweenLines spacingBetweenLines = new Wp.SpacingBetweenLines();
            spacingBetweenLines.LineRule = Wp.LineSpacingRuleValues.Exact;
            spacingBetweenLines.Line = "200";
            //paragraphProperties.Append(spacingBetweenLines); It seems like only the first cell gets this
            paragraph.Append(paragraphProperties);
            Wp.Run run = new Wp.Run();
            run.Append(table);
            paragraph.Append(run);
            return paragraph;
        }
        public static Wp.Paragraph GenerateWelcomeParagraph(String danceName, String danceDates)
        {
            Wp.Paragraph paragraph1 = new Wp.Paragraph();

            String[] lines = { myTexts["welcome"], myTexts["to"], danceName, danceDates };

            String[] colors = { "Black", "Black", "Black", "Black" };
            int[] fontSizes = { 20, 12, 32, 20 };
            return GenerateParagraph(lines, fontSizes, colors);
        }
        public static Wp.Paragraph GenerateCallerNameParagraph(String callerName)
        {
            String[] names = callerName.Split('_');

            String firstName = char.ToUpper(names[0][0]) + names[0].Substring(1);
            String lastName = char.ToUpper(names[1][0]) + names[1].Substring(1);

            callerName = String.Format("{0} {1}", firstName, lastName);
            String[] lines = { callerName };
            String[] colors = { "Black" };
            int[] fontSizes = { 32 };
            return GenerateParagraph(lines, fontSizes, colors);
        }

        public static Wp.Paragraph GenerateFeesParagraph(SchemaInfo schemaInfo)
        {
            List<String> lines = new List<String>();
            if (schemaInfo.schemaName.StartsWith("weekend"))
            {
                String line1 = String.Format(myTexts["weekend_member_fees"], myFees.weekends[0], myFees.weekends[1]);
                htmlStringBuilder.Append("<p class='m8_schema'>\n");
                lines.Add(line1);
                htmlStringBuilder.Append(line1+"<br>");
                String line2 = String.Format(myTexts["weekend_non_member_fees"], myFees.weekends[2], myFees.weekends[3]);
                lines.Add(line2);
                htmlStringBuilder.Append(line2+"<br>");
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
            return GenerateParagraph(lines.ToArray(), fontSizes, colors);
        }
        public static Wp.Paragraph GenerateDanceLocationParagraph(String danceLocation, int maxWidth)
        {
            String[] lines = { danceLocation };
            int[] fontSizes = { 14 };
            String[] colors = { "Black" };
            //224/197/18
            Wp.ParagraphBorders borders = createParagraphBorders(Wp.BorderValues.Double, 12, "E0C512");
            htmlStringBuilder.Append(String.Format("<p  class='m8_schema m8_border'  style='max-width: {0}px;'>{1}</p>",
                maxWidth, danceLocation));
            return GenerateParagraph(lines, fontSizes, colors, borders);
        }
        public static Wp.Paragraph GenerateCoffeeParagraph(Boolean coffee, int maxWidth)
        {
            String text = String.Format("{0}-{1}", myTexts["no_coffee"], myTexts["lunch"]);
            if (coffee)
            {
                text = String.Format("{0}-{1}", myTexts["coffee"], myTexts["lunch"]);
            }
            String[] lines = { text };
            int[] fontSizes = { 13 };
            String[] colors = { "Black" };
            Wp.ParagraphBorders borders = createParagraphBorders(Wp.BorderValues.Double, 12, "E0C512");
            htmlStringBuilder.Append(String.Format("<p  class='m8_schema m8_border' style='max-width: {0}px;'>{1}</p>",
                maxWidth, text));

            return GenerateParagraph(lines, fontSizes, colors,borders);
        }

        public static Wp.Paragraph GenerateRotationParagraph()
        {
            String[] lines = { myTexts["rotation"] };
            int[] fontSizes = { 14 };
            String[] colors = { "Black" };
            htmlStringBuilder.Append("<p  class='m8_schema'>\n");
            htmlStringBuilder.Append(lines[0] + "</p>\n");
            return GenerateParagraph(lines, fontSizes, colors);
        }
        private static Wp.Paragraph GenerateParagraph(string[] lines, int[] fontSizes, String[] colors, Wp.ParagraphBorders borders = null)
        {
            Wp.Paragraph paragraph = new Wp.Paragraph();


            for (int i = 0; i < lines.Length; i++)
            {
                String fontSizeTxt = (fontSizes[i] * 2).ToString();


                //byte[] bytes= Encoding.Default.GetBytes(lines[i]);
                //String line = Encoding.UTF8.GetString(bytes); 
                String line = lines[i];


                Wp.FontSize fontSize = new Wp.FontSize { Val = new OXML.StringValue(fontSizeTxt) };  // Size in half points
                //Wp.ParagraphProperties paragraphProperties = new Wp.ParagraphProperties();
                Wp.ParagraphProperties paragraphProperties = new Wp.ParagraphProperties(borders);
                Wp.Justification justification = new Wp.Justification() { Val = Wp.JustificationValues.Center };
                paragraphProperties.Append(justification);

                Wp.Run run = new Wp.Run();
                Wp.RunProperties runProperties = new Wp.RunProperties();
                Wp.RunFonts runFonts = new Wp.RunFonts { Ascii = "Comic Sans MS", HighAnsi = "Comic Sans MS" };
                runProperties.Append(runFonts);
                runProperties.Append(fontSize);
                Wp.Text text1 = new Wp.Text();
                text1.Text = line;
                if (colors[i] == "Red")
                {
                    Wp.Color color = new Wp.Color() { Val = "FF0000" };
                    runProperties.Append(color);

                }
                run.Append(runProperties);
                run.Append(text1);
                if (i < lines.Length-1)
                {
                    run.Append(new Wp.Break());
                }
                paragraph.Append(paragraphProperties);
                paragraph.Append(run);

            }
            return paragraph;
        }

        // Insert a table into a word processing document.
        public static Wp.Table CreateDanceSchemaTable(String lang, SchemaInfo schemaInfo, String schemaName)
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
                return createFestivalDanceSchemaTable(schemaInfo);
            }
            else
            {
                return null;
            }
        }
        public static Wp.Table createFestivalDanceSchemaTable(SchemaInfo schemaInfo)
        {
            Wp.Table table = new Wp.Table();
            Wp.TableProperties tblProp = new Wp.TableProperties(createTableBorders(Wp.BorderValues.Dashed, 12));
            return table;
        }
        public static Wp.Table createWeekendDanceSchemaTable(String lang, List<DancePass[]> dancePassesDayList, SchemaInfo schemaInfo)
        {
            //}
            Wp.Table table = new Wp.Table();


            int[] colWidth = { 2000, 500, 300, 2000, 500 };

            createFirstWeekendRow(table, colWidth);                                 // row 1, Header row
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

        private static void createWeekEndRow(List<DancePass[]> dancePassesDayList, int rowNumber, out string level2, out string timeString2, out string level1, out string timeString1)
        {
            int i1 = rowNumber-2;
            level2 = "";
            timeString2 = "";
            try
            {
                timeString2 = formatTimeInterval(dancePassesDayList[1], i1);
                level2 = dancePassesDayList[1][i1].level;
            }
            catch (Exception e)
            {

            }
            level1 = "";
            timeString1 = "";
            try
            {
                timeString1 = formatTimeInterval(dancePassesDayList[0], i1);
                level1 = dancePassesDayList[0][i1].level;
            }
            catch (Exception e)
            {

            }
        }

        private static String createWeekEndRowHtml(String lang, List<DancePass[]> dancePassesDayList, int rowNumber, SchemaInfo schemaInfo, Boolean merge = false)
        {
            string level2, timeString2, level1, timeString1;
            createWeekEndRow(dancePassesDayList, rowNumber, out level2, out timeString2, out level1, out timeString1);


           String row= String.Format("<tr class='m8_schema'><td class='m8_schema m8_time'>{0}</td><td class='m8_schema m8_level'>{1}</td><td class='m8_schema m8_space'> </td><td class='m8_schema m8_time'>{2}</td><td class='m8_schema m8_level'>{3}</td></tr>",
               timeString1, level1, timeString2, level2);
            return row;
            
        }



        /// <summary>
        /// Create the header row for a Weekend dance schedule
        /// </summary>
        /// <param name="table"></param>
        /// <param name="colWidth"></param>
        private static void createFirstWeekendRow(Wp.Table table, int[] colWidth)
        {
            String[] content = new String[5];// = { "Lördag", "", "Söndag", "" };
            content[0] = myTexts["Saturday"];
            content[3] = myTexts["Sunday"];

            Wp.TableRow tr1 = createRow1(content, colWidth);
            table.Append(tr1);
        }

        private static String formatTimeInterval(DancePass[] dancePasses, int row)
        {
            String text = String.Format("{0}-{1}", dancePasses[row].start_time, dancePasses[row].end_time);
            if (dancePasses[row].end_time.Length > 6) // quick and dirty solution for Årsmöte
            {
                text = String.Format("{0} {1}", dancePasses[row].start_time, dancePasses[row].end_time);
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

        private static Wp.TableRow createRow1(String[] content, int[] colWidth)
        {
            Wp.TableRow row = new Wp.TableRow();
            var rowProps = new Wp.TableRowProperties();


            rowProps.Append(new Wp.TableJustification { Val = Wp.TableRowAlignmentValues.Center });

            row.Append(rowProps);
            Wp.TableCell[] tableCell = new Wp.TableCell[5];
            for (int i = 0; i < content.Length; i++)
            {
                tableCell[i] = createACell(content[i], colWidth[i], true);
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
        private static Wp.TableRow createRow(String[] content, int[] colWidth, bool merge = false)
        {

            Wp.TableRow row = new Wp.TableRow();
            var rowProps = new Wp.TableRowProperties();
            rowProps.Append(new Wp.TableJustification { Val = Wp.TableRowAlignmentValues.Center });

            row.Append(rowProps);
            List<Wp.TableCell> tableCells = new List<Wp.TableCell>();
            for (int i = 0; i < content.Length; i++)
            {
                Wp.TableCell tableCell1 = createACell(content[i], colWidth[i]);
                tableCells.Add(tableCell1);
                row.Append(tableCell1);
            }
            if (merge)
            {
                MergeCells(tableCells[0], tableCells[1]);
            }
            return row;

        }
        private static Wp.TableCell createACell(String text, int width, Boolean underline = false)
        {
            String widthStr = width.ToString();
            Wp.TableCell tableCell = new Wp.TableCell();
            Wp.FontSize fontSize = new Wp.FontSize { Val = "28" };  // Size in half points
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
            runProperties.Append(fontSize);
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

        private static OXML.OpenXmlElement[] createBorders(Wp.BorderValues type, OXML.UInt32Value size, String color="000000")
        {
            OXML.OpenXmlElement[] borders = {
            new Wp.TopBorder()
                {
                    Val = new OXML.EnumValue<Wp.BorderValues>(type),
                    Size = size,
                    Color= color
            },
            new Wp.BottomBorder()
                {
                    Val = new OXML.EnumValue<Wp.BorderValues>(type),
                    Size = size,
                    Color= color

                },
            new Wp.LeftBorder()
                {
                    Val = new OXML.EnumValue<Wp.BorderValues>(type),
                    Size = size,
                    Color= color

                },
            new Wp.RightBorder()
                {
                    Val = new OXML.EnumValue<Wp.BorderValues>(type),
                    Size = size,
                    Color= color

                },
            new Wp.InsideHorizontalBorder()
                {
                    Val = new OXML.EnumValue<Wp.BorderValues>(type),
                    Size = size,
                    Color= color

                },
            new Wp.InsideVerticalBorder()
                {
                    Val = new OXML.EnumValue<Wp.BorderValues>(type),
                    Size = size,
                    Color= color

                }
            };
            return borders;
        }
        public static Wp.ParagraphBorders createParagraphBorders(Wp.BorderValues type, OXML.UInt32Value size, String color="000000")
        {
            Wp.ParagraphBorders paragraphBorders = new Wp.ParagraphBorders(createBorders(type, size, color));
            return paragraphBorders;
        }
        public static Wp.TableBorders createTableBorders(Wp.BorderValues type, OXML.UInt32Value size)
        {
            Wp.TableBorders tableBorders = new Wp.TableBorders(createBorders(type, size));
            return tableBorders;
        }
        public static void addImage(String type, WordprocessingDocument wordprocessingDocument, String fileName, int maxHeight, double x0_mm, double y0_mm)
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
                        scale = (double) maxHeight / (double) iHeight;
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
                AddInlineImageToBody(wordprocessingDocument, mainPart.GetIdOfPart(imagePart), (int)(scale * iWidth), (int)(scale * iHeight));
            }
        }
        public static void AddAnchorImageToBody(WordprocessingDocument wordprocessingDocument, String relationshipId, int iWidth, int iHeight, double x0_mm, double y0_mm)
        {
            // Define the reference of the image.
            //var x = new DW.Anchor();
            var element = GetAnchorPicture(relationshipId, x0_mm, y0_mm, iWidth, iHeight);
            // Append the reference to the body. The element should be in 
            // a DocumentFormat.OpenXml.Wordprocessing.Run.
            wordprocessingDocument.MainDocumentPart.Document.Body.AppendChild(new Wp.Paragraph(new Wp.Run(element)));

        }

        public static void AddInlineImageToBody(WordprocessingDocument wordprocessingDocument, String relationshipId, int iWidth, int iHeight)
        {
            // convert the pixels to EMUs this way:
            iWidth = (int)Math.Round((decimal)iWidth * 9525);
            iHeight = (int)Math.Round((decimal)iHeight * 9525);
            // Define the reference of the image.
            var element =
                 new Wp.Drawing(
                     new DW.Inline(
                         new DW.Extent() { Cx = iWidth, Cy = iHeight },
                         new DW.EffectExtent()
                         {
                             LeftEdge = 0L,
                             TopEdge = 0L,
                             RightEdge = 0L,
                             BottomEdge = 0L
                         },
                         new DW.DocProperties()
                         {
                             Id = (OXML.UInt32Value)1U,
                             Name = "Picture 1"
                         },
                         new DW.NonVisualGraphicFrameDrawingProperties(
                             new A.GraphicFrameLocks() { NoChangeAspect = true }),
                         new A.Graphic(
                             new A.GraphicData(
                                 new PIC.Picture(
                                     new PIC.NonVisualPictureProperties(
                                         new PIC.NonVisualDrawingProperties()
                                         {
                                             Id = (OXML.UInt32Value)0U,
                                             Name = "M8-logo1.gif"
                                         },
                                         new PIC.NonVisualPictureDrawingProperties()),
                                     new PIC.BlipFill(
                                         new A.Blip(
                                             new A.BlipExtensionList(
                                                 new A.BlipExtension()
                                                 {
                                                     Uri =
                                                       "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                                 })
                                         )
                                         {
                                             Embed = relationshipId,
                                             CompressionState =
                                             A.BlipCompressionValues.Print
                                         },
                                         new A.Stretch(
                                             new A.FillRectangle())),
                                     new PIC.ShapeProperties(
                                         new A.Transform2D(
                                             new A.Offset() { X = 0L, Y = 0L },
                                             new A.Extents() { Cx = iWidth, Cy = iHeight }),
                                         new A.PresetGeometry(
                                             new A.AdjustValueList()
                                         )
                                         { Preset = A.ShapeTypeValues.Rectangle }))
                             )
                             { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                     )
                     {
                         DistanceFromTop = (OXML.UInt32Value)1000U,
                         DistanceFromBottom = (OXML.UInt32Value)0U,
                         DistanceFromLeft = (OXML.UInt32Value)0U,
                         DistanceFromRight = (OXML.UInt32Value)0U,
                         EditId = "50D07946"
                     });

            // Append the reference to the body. The element should be in 
            // a DocumentFormat.OpenXml.Wordprocessing.Run.
            //wordprocessingDocument.MainDocumentPart.Document.Body.AppendChild(new Wp.Paragraph(new Wp.Run(element)));
            wordprocessingDocument.MainDocumentPart.Document.Body.AppendChild(
                new Wp.Paragraph(new Wp.Run(element))
                {
                    ParagraphProperties = new Wp.ParagraphProperties()
                    {
                        Justification = new Wp.Justification()
                        {
                            Val = Wp.JustificationValues.Center
                        }
                    }
                });
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="imagePartId"></param>
        /// <param name="x0_mm">Logo distance from top of page, cm</param>
        /// <param name="y0_mm">Logo distance from left edge of page, cm</param>
        /// <param name="wPixels">width of logo picture</param>
        /// <param name="hPixels">height of logo picture</param>
        /// <returns></returns>
        public static Wp.Drawing GetAnchorPicture(String imagePartId, double x0_mm, double y0_mm, int wPixels, int hPixels)
        {
            long iWidth = (long)Math.Round((decimal)wPixels * 9525);
            long iHeight = (long)Math.Round((decimal)hPixels * 9525);
            long mmToEmu = 36000L;
            Wp.Drawing _drawing = new Wp.Drawing();
            DW.Anchor _anchor = new DW.Anchor()
            {
                DistanceFromTop = (OXML.UInt32Value)0U,
                DistanceFromBottom = (OXML.UInt32Value)0U,
                SimplePos = false,
                RelativeHeight = (OXML.UInt32Value)251658240U,
                BehindDoc = true,
                Locked = false,
                LayoutInCell = true,
                AllowOverlap = true,
            };

            DW.SimplePosition _spos = new DW.SimplePosition()
            {
                X = 0,
                Y = 0
            };

            DW.HorizontalPosition _hp = new DW.HorizontalPosition()
            {
                RelativeFrom = DW.HorizontalRelativePositionValues.LeftMargin
            };
            DW.PositionOffset _hPO = new DW.PositionOffset();
            _hPO.Text = (x0_mm* mmToEmu).ToString();  // Convert mm to EMU
            _hp.Append(_hPO);

            DW.VerticalPosition _vp = new DW.VerticalPosition()
            {
                RelativeFrom = DW.VerticalRelativePositionValues.TopMargin
            };
            DW.PositionOffset _vPO = new DW.PositionOffset();
            _vPO.Text = (y0_mm * mmToEmu).ToString();   // Convert mm to EMU
            _vp.Append(_vPO);

            DW.Extent _e = new DW.Extent()
            {
                //Cx = 989965L,
                //Cy = 791845L
                Cx = iWidth,
                Cy = iHeight
            };

            DW.EffectExtent _ee = new DW.EffectExtent()
            {
                LeftEdge = 0L,
                TopEdge = 0L,
                RightEdge = 0L,
                BottomEdge = 0L
            };
            DW.WrapNone _wpn = new DW.WrapNone();

            DW.DocProperties _dp = new DW.DocProperties()
            {
                Id = 1U,
                Name = "Test Picture"
            };

            OXML.Drawing.Graphic _g = new OXML.Drawing.Graphic();
            OXML.Drawing.GraphicData _gd = new OXML.Drawing.GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" };
            OXML.Drawing.Pictures.Picture _pic = new OXML.Drawing.Pictures.Picture();

            OXML.Drawing.Pictures.NonVisualPictureProperties _nvpp = new OXML.Drawing.Pictures.NonVisualPictureProperties();
            OXML.Drawing.Pictures.NonVisualDrawingProperties _nvdp = new OXML.Drawing.Pictures.NonVisualDrawingProperties()
            {
                Id = 0,
                Name = "Picture"
            };
            OXML.Drawing.Pictures.NonVisualPictureDrawingProperties _nvpdp = new OXML.Drawing.Pictures.NonVisualPictureDrawingProperties();
            _nvpp.Append(_nvdp);
            _nvpp.Append(_nvpdp);


            OXML.Drawing.Pictures.BlipFill _bf = new OXML.Drawing.Pictures.BlipFill();
            OXML.Drawing.Blip _b = new OXML.Drawing.Blip()
            {
                Embed = imagePartId,
                CompressionState = OXML.Drawing.BlipCompressionValues.Print
            };
            _bf.Append(_b);

            OXML.Drawing.Stretch _str = new OXML.Drawing.Stretch();
            OXML.Drawing.FillRectangle _fr = new OXML.Drawing.FillRectangle();
            _str.Append(_fr);
            _bf.Append(_str);

            OXML.Drawing.Pictures.ShapeProperties _shp = new OXML.Drawing.Pictures.ShapeProperties();
            OXML.Drawing.Transform2D _t2d = new OXML.Drawing.Transform2D();
            OXML.Drawing.Offset _os = new OXML.Drawing.Offset()
            {
                X = 0L,
                Y = 0L
            };
            OXML.Drawing.Extents _ex = new OXML.Drawing.Extents()
            {
                //Cx = 989965L,
                //Cy = 791845L
                Cx = iWidth,
                Cy = iHeight
            };

            _t2d.Append(_os);
            _t2d.Append(_ex);

            OXML.Drawing.PresetGeometry _preGeo = new OXML.Drawing.PresetGeometry()
            {
                Preset = OXML.Drawing.ShapeTypeValues.Rectangle
            };
            OXML.Drawing.AdjustValueList _adl = new OXML.Drawing.AdjustValueList();
            _preGeo.Append(_adl);

            _shp.Append(_t2d);
            _shp.Append(_preGeo);

            _pic.Append(_nvpp);
            _pic.Append(_bf);
            _pic.Append(_shp);

            _gd.Append(_pic);
            _g.Append(_gd);

            _anchor.Append(_spos);
            _anchor.Append(_hp);
            _anchor.Append(_vp);
            _anchor.Append(_e);
            _anchor.Append(_ee);
            //_anchor.Append(_wp);
            _anchor.Append(_wpn);
            _anchor.Append(_dp);
            _anchor.Append(_g);

            _drawing.Append(_anchor);

            return _drawing;
        }


    }
}
