using System;
using System.IO;
using System.Drawing;
using System.Globalization;
using System.Collections.Generic;
using System.Net.Cache;
using System.Linq;

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
        private static Wp.Color wpColorBlackx = new Wp.Color() { Val = "000000" };
        private static Wp.Color wpColorRedx = new Wp.Color() { Val = "FF0000" };
        private static Dictionary<String, String> texts;
        private static List<DancePass[]> dancePassesDayList = new List<DancePass[]>();

        /// <summary>
        /// Only weekend dances supported 
        /// </summary>
        /// <param name="filepath"></param>
        /// <param name="danceName"></param>
        /// <param name="danceDateStart"></param>
        /// <param name="danceDateEnd"></param>
        public static void CreateWordprocessingDocument(Dictionary<String, String> myTexts, String lang, SchemaInfo schemaInfo,
            String schemaName, String path, DateTime danceDateStart, DateTime danceDateEnd)
        {
            string monthName1, monthName2, danceDates;

            texts = myTexts;
            danceDates = createDanceDates(danceDateStart, danceDateEnd, out monthName1, out monthName2);

            //String htmlText = GetHtmlCode(danceDateStart, danceDateEnd, monthName1, monthName2);
            //File.WriteAllText(Path.Combine(texts["outputFolder"], String.Format("{0}.html", texts["danceName"])), htmlText);
            //// Create a document by supplying the filepath. 

            try
            {
                using (WordprocessingDocument wordDocument =
                    WordprocessingDocument.Create(path, OXML.WordprocessingDocumentType.Document))
                {
                    // Add a main document part. 
                    MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();

                    mainPart.Document = new Wp.Document();
                    Wp.Body body = mainPart.Document.AppendChild(new Wp.Body());

                    //String fileNameLogo = @"Resources\M8-logo1.gif";
                    String fileNameLogo = "https://motiv8s.se/19/images/M8/Logga_Transparent.jpg";
                    addImage("Anchor", wordDocument, fileNameLogo, 0.1667, 1.0, 1.6);
                    Wp.Paragraph paragraph1 = GenerateWelcomeParagraph(texts["danceName"].ToUpper(), danceDates);
                    body.AppendChild(paragraph1);

                    addImage("Inline", wordDocument, texts["callerPictureFile"], 0.7, 6.0, 10.0);

                    body.AppendChild(GenerateCallerNameParagraph(texts["callerName"]));

                    Wp.Table table = CreateDanceSchemaTable(lang, schemaInfo, schemaName);
                    Wp.Paragraph tableParagraph = generateTableParagraph(table);
                    //body.AppendChild(table);
                    body.AppendChild(tableParagraph);


                    //body.AppendChild(new Wp.Break());
                    body.AppendChild(GenerateParagraph4());
                    body.AppendChild(GenerateParagraph5());
                    body.AppendChild(GenerateParagraph6());
                }
            }
            catch (Exception e)
            {
                throw e;
            }
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
            String danceDates = String.Format("{0} - {1} {2}", danceDateStart.Day, danceDateEnd.Day, monthName1);
            if (month1 != month2)
            {
                danceDates = String.Format("{0}{1} - {2} {3}", danceDateStart.Day, danceDateEnd.Day, monthName1, monthName2);
            }
            return danceDates;
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
            //Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "004F7104", RsidParagraphProperties = "008F2986", RsidRunAdditionDefault = "004F7104" };
            Wp.Paragraph paragraph1 = new Wp.Paragraph();

            String[] lines = { texts["welcome"], texts["to"], danceName, danceDates };

            String[] colors = { "Black", "Black", "Black", "Black" };
            int[] fontSizes = { 20, 12, 32, 20 };
            return GenerateParagraph(lines, fontSizes, colors);
        }
        public static Wp.Paragraph GenerateCallerNameParagraph(String callerName)
        {
            String[] names = callerName.Split('_');

            String firstName = char.ToUpper(names[0][0]) + names[0].Substring(1);
            String lastName = char.ToUpper(names[1][0]) + names[1].Substring(1);

            callerName = String.Format("{0} {1}", names[0], names[1]);
            String[] lines = { callerName };
            String[] colors = { "Black" };
            int[] fontSizes = { 32 };
            return GenerateParagraph(lines, fontSizes, colors);
        }



        public static Wp.Paragraph GenerateParagraph4()
        {
            String[] lines = {
                "Medlem: 100 kr/pass, samtliga pass 350 kr",
                "Ej medlem: 120 kr/pass, samtliga pass 400 kr",
                "Betala gärna i förväg på PlusGiro 85 56 69-8 (MOTIV8'S)",
                "Swish till 070-422 82 27 (Arne G) eller kontanter ”i dörren” går också bra",
            };
            int[] fontSizes = { 12, 12, 12, 12 };
            String[] colors = { "Black", "Black", "Black", "Black" };
            return GenerateParagraph(lines, fontSizes, colors);
        }
        public static Wp.Paragraph GenerateParagraph5()
        {
            String[] lines = {
                "Plats: Segersjö Folkets Hus, Scheelevägen 41 i Tumba"
            };
            int[] fontSizes = { 14 };
            String[] colors = { "Red" };
            return GenerateParagraph(lines, fontSizes, colors);
        }
        public static Wp.Paragraph GenerateParagraph6()
        {
            String[] lines = {
                "Vi använder rotationsprogram på samtliga nivåer!",
                "Ta med eget fika! - Vi ordnar hämtning av Pizza och sallad till pausen!"
            };
            int[] fontSizes = { 14, 14 };
            String[] colors = { "Black", "Black" };

            return GenerateParagraph(lines, fontSizes, colors);
        }

        private static Wp.Paragraph GenerateParagraph(string[] lines, int[] fontSizes, String[] colors)
        {
            Wp.Paragraph paragraph = new Wp.Paragraph();
            for (int i = 0; i < lines.Length; i++)
            {
                String fontSizeTxt = (fontSizes[i] * 2).ToString();


                //byte[] bytes= Encoding.Default.GetBytes(lines[i]);
                //String line = Encoding.UTF8.GetString(bytes); 
                String line = lines[i];


                Wp.FontSize fontSize = new Wp.FontSize { Val = new OXML.StringValue(fontSizeTxt) };  // Size in half points
                Wp.ParagraphProperties paragraphProperties = new Wp.ParagraphProperties();
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
                run.Append(new Wp.Break());
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
                return createWeekendDanceSchemaTable(lang, dancePassesDayList, schemaInfo);
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
            return table;
        }
        public static Wp.Table createWeekendDanceSchemaTable(String lang, List<DancePass[]> dancePassesDayList, SchemaInfo schemaInfo)
        {
            //}
            Wp.Table table = new Wp.Table();

            Wp.TableProperties tblProp = new Wp.TableProperties(createTableBorders(Wp.BorderValues.Dashed, 12));

            int[] colWidth = { 2000, 500, 300, 2000, 500 };

            createFirstWeekendRow(table, colWidth);                                 // row 1, Header row
            createWeekEndRow(dancePassesDayList, table, schemaInfo.colWidth, 2);    // row 2
            // Merge column 1 and 2 if schemaName== "weekend_meeting"
            createWeekEndRow(dancePassesDayList, table, schemaInfo.colWidth, 3, schemaInfo.schemaName == "weekend_meeting");   // row 3
            if (dancePassesDayList[0].Length > 2 || dancePassesDayList[1].Length > 2)
            {
                createWeekEndRow(dancePassesDayList, table, schemaInfo.colWidth, 4);  // row 4
            }

            return table;

        }

        private static void createWeekEndRow(List<DancePass[]> dancePassesDayList, Wp.Table table, List<int> colWidth, int row, Boolean merge = false)
        {
            int i1 = row - 2;
            String level2 = "";
            String timeString2 = "";
            try
            {
                timeString2 = formatTimeInterval(dancePassesDayList[1], i1);
                level2 = dancePassesDayList[1][i1].level;
            }
            catch (Exception e)
            {

            }
            String level1 = "";
            String timeString1 = "";
            try
            {
                timeString1 = formatTimeInterval(dancePassesDayList[0], i1);
                level1 = dancePassesDayList[0][i1].level;
            }
            catch (Exception e)
            {

            }


            String[] content = { timeString1,
                level1,
                "",
                timeString2,
                level2 };

            table.Append(createRow(content, colWidth.ToArray(), merge));
        }

        /// <summary>
        /// Create the header row for a Weekend dance schedule
        /// </summary>
        /// <param name="table"></param>
        /// <param name="colWidth"></param>
        private static void createFirstWeekendRow(Wp.Table table, int[] colWidth)
        {
            String[] content = new String[5];// = { "Lördag", "", "Söndag", "" };
            content[0] = texts["Saturday"];
            content[3] = texts["Sunday"];

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
        public static Wp.TableBorders createTableBorders(Wp.BorderValues type, OXML.UInt32Value size)
        {
            Wp.TableBorders tableBorders = new Wp.TableBorders(
                new Wp.TopBorder()
                {
                    Val = new OXML.EnumValue<Wp.BorderValues>(type),
                    Size = size
                },
                new Wp.BottomBorder()
                {
                    Val = new OXML.EnumValue<Wp.BorderValues>(type),
                    Size = size
                },
                new Wp.LeftBorder()
                {
                    Val = new OXML.EnumValue<Wp.BorderValues>(type),
                    Size = size
                },
                new Wp.RightBorder()
                {
                    Val = new OXML.EnumValue<Wp.BorderValues>(type),
                    Size = size
                },
                new Wp.InsideHorizontalBorder()
                {
                    Val = new OXML.EnumValue<Wp.BorderValues>(type),
                    Size = size
                },
                new Wp.InsideVerticalBorder()
                {
                    Val = new OXML.EnumValue<Wp.BorderValues>(type),
                    Size = size
                }
            );
            return tableBorders;
        }
        public static void addImage(String type, WordprocessingDocument wordprocessingDocument, String fileName, double scale, double x0_cm, double y0_cm)
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



            System.Net.WebRequest request = System.Net.WebRequest.Create(fileName);
            request.CachePolicy = noCachePolicy;
            System.Net.WebResponse response = request.GetResponse();
            using (System.IO.Stream responseStream = response.GetResponseStream())
            {
                Console.WriteLine("IsFromCache? {0}", response.IsFromCache);
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


                    memoryStream.Position = 0;
                    imagePart.FeedData(memoryStream);
                }
            }
            if (type == "Anchor")
            {
                AddAnchorImageToBody(wordprocessingDocument, mainPart.GetIdOfPart(imagePart), (int)(scale * iWidth), (int)(scale * iHeight), x0_cm, y0_cm);
            }
            else
            {
                AddInlineImageToBody(wordprocessingDocument, mainPart.GetIdOfPart(imagePart), (int)(scale * iWidth), (int)(scale * iHeight));
            }
        }
        public static void AddAnchorImageToBody(WordprocessingDocument wordprocessingDocument, String relationshipId, int iWidth, int iHeight, double x0_cm, double y0_cm)
        {
            // Define the reference of the image.
            //var x = new DW.Anchor();
            var element = GetAnchorPicture(relationshipId, x0_cm, y0_cm, iWidth, iHeight);
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
        /// <param name="x0_cm">Logo distance from top of page, cm</param>
        /// <param name="y0_cm">Logo distance from left edge of page, cm</param>
        /// <param name="wPixels">width of logo picture</param>
        /// <param name="hPixels">height of logo picture</param>
        /// <returns></returns>
        public static Wp.Drawing GetAnchorPicture(String imagePartId, double x0_cm, double y0_cm, int wPixels, int hPixels)
        {
            // convert the cm to EMUs this way:
            long f1 = 360000;
            //long x0_emu = (long)Math.Round(x0_cm * f1);
            //long y0_emu = (long)Math.Round(y0_cm * f1);
            // convert the pixels to EMUs this way:
            long iWidth = (long)Math.Round((decimal)wPixels * 9525);
            long iHeight = (long)Math.Round((decimal)hPixels * 9525);

            Wp.Drawing _drawing = new Wp.Drawing();
            DW.Anchor _anchor = new DW.Anchor()
            {
                DistanceFromTop = (OXML.UInt32Value)0U,
                DistanceFromBottom = (OXML.UInt32Value)0U,
                DistanceFromLeft = (OXML.UInt32Value)114300U,
                DistanceFromRight = (OXML.UInt32Value)114300U,
                SimplePos = false,
                RelativeHeight = (OXML.UInt32Value)251658240U,
                BehindDoc = true,
                Locked = false,
                LayoutInCell = true,
                AllowOverlap = true,
                EditId = "44CEF5E4",
                AnchorId = "44803ED1"
            };

            DW.SimplePosition _spos = new DW.SimplePosition()
            {
                X = 0,
                Y = 0
            };

            DW.HorizontalPosition _hp = new DW.HorizontalPosition()
            {
                RelativeFrom = DW.HorizontalRelativePositionValues.Column
            };
            DW.PositionOffset _hPO = new DW.PositionOffset();
            _hPO.Text = "4445";
            _hp.Append(_hPO);

            DW.VerticalPosition _vp = new DW.VerticalPosition()
            {
                RelativeFrom = DW.VerticalRelativePositionValues.Paragraph
            };
            DW.PositionOffset _vPO = new DW.PositionOffset();
            _vPO.Text = "0";
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

        public static String GetHtmlCode(DateTime danceDateStart, DateTime danceDateEnd, String monthName1, String monthName2)
        {

            var sb = new System.Text.StringBuilder();
            sb.Append(String.Format("<p>{0} {1} {2}<br/>{3} - {4}<br/>{5} - {6} </p>", texts["Saturday"], danceDateStart.Day, monthName1,
                texts["pass_1_weekend_time"], texts["pass_1_weekend_level"], texts["pass_2_weekend_time"], texts["pass_2_weekend_level"]));
            sb.Append(String.Format("<p>Söndag {0} {1} <br/>10:00 - 13:00 - C3A <br/> 14:00 - 17:00 - C3B</p>", danceDateEnd.Day, monthName2));
            sb.Append("<p class='mobile - undersized - lower'>Medlem: <span style='background - color: #ffff00;'>100</span> kr/pass, samtliga pass <span style='background-color: #ffff00;'>350</span> kr<br/>");
            sb.Append("Ej medlem: <span style='background - color: #ffff00;'>120</span> kr/pass, samtliga pass <span style='background-color: #ffff00;'>400</span> kr<br/>");
            sb.Append("Betala gärna i förväg på PlusGiro 85 56 69-8 (MOTIV8'S)<br/> Swish till 070-422 82 27 (Arne G) eller kontanter ”i dörren” går också bra.</p>");
            sb.Append("<p><span style='color: #ff0000;'>OBS! Plats: Segersjö Folkets Hus, Scheelevägen 41 i Tumba</span></p>");
            sb.Append("<ul>");
            sb.Append("<li>Vi använder rotationsprogram på samtliga nivåer!</li>");
            sb.Append("<li>Ta med eget fika!</li>");
            sb.Append("<li>Vi ordnar hämtning av Pizza och sallad till pausen.</li>");
            sb.Append("</ul>");
            sb.Append("<p><em><em>Om du har några symtom som halsont, snuva, feber, hosta eller sjukdomskänsla så ska du stanna hemma.<br />Reservation för att vi inte kan genomföra dansen på grund av restriktioner.</em></em></p>");
            return sb.ToString();
        }
        /*







         
        */
    }
}
