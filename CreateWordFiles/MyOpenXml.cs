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
    internal class MyOpenXml
    {
        public static String Motiv8sColor = "E0C512";
        /// <summary>
        /// 
        /// </summary>
        /// <param name="imagePartId"></param>
        /// <param name="x0_mm">Logo distance from top of page, cm</param>
        /// <param name="y0_mm">Logo distance from left edge of page, cm</param>
        /// <param name="wPixels">width of logo picture</param>
        /// <param name="hPixels">height of logo picture</param>
        /// <returns></returns>

        public static void AddInlineImageToBody(WordprocessingDocument wordprocessingDocument, String relationshipId,
            int iWidth, int iHeight, Boolean useBorder, String fileName)
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
                                             Name = fileName
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
            Wp.ParagraphBorders borders = MyOpenXml.createParagraphBorders(Wp.BorderValues.None, 0);
            if (useBorder)
            {
                borders = MyOpenXml.createParagraphBorders(Wp.BorderValues.Double, 12);
            }

            // Append the reference to the body. The element should be in 
            // a DocumentFormat.OpenXml.Wordprocessing.Run.
            //wordprocessingDocument.MainDocumentPart.Document.Body.AppendChild(new Wp.Paragraph(new Wp.Run(element)));
            Wp.Run run = new Wp.Run(element);
            //Wp.Border border = new Wp.Border() { Val = Wp.BorderValues.Single };
            Wp.RunProperties runProp= new Wp.RunProperties();
            //border.Color = "FF0000";
            //border.Size = 3 * 4;  // Size in pixels/4
            
           // runProp.Append(createBorders(Wp.BorderValues.Double, 12));

            run.Append(runProp);
            wordprocessingDocument.MainDocumentPart.Document.Body.AppendChild(
                //  new Wp.Paragraph(new Wp.Run(element))
                new Wp.Paragraph(run)
                {
                    ParagraphProperties = new Wp.ParagraphProperties()
                    {
                        Justification = new Wp.Justification()
                        {
                            Val = Wp.JustificationValues.Center
                        },
                        ParagraphBorders = borders
                    }
                }) ;
        }


        private static OXML.OpenXmlElement[] createBorders(Wp.BorderValues type, OXML.UInt32Value size,
            Wp.BorderValues borderTypeOuter = Wp.BorderValues.Single,
            Wp.BorderValues borderTypeInner = Wp.BorderValues.None, String color = "E0C512")
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

        public static Wp.ParagraphBorders createParagraphBorders(Wp.BorderValues type, OXML.UInt32Value size)
        {
            Wp.ParagraphBorders paragraphBorders = new Wp.ParagraphBorders(createBorders(type, size));
            return paragraphBorders;
        }


        /// <summary>
        /// Adds a footer to a document
        /// </summary>
        /// <param name="document"></param>
        /// <param name="footerText"></param>
        /// <param name="fontSize">Font size in points</param>
        public static void SetMarginsAndFooter(WordprocessingDocument document, String footerText, int fontSize, int topMargin = 200, uint leftMargin = 1008, uint rigthMargin = 1008, int bottomMargin = 1008)
        {
            // Get the main document part.
            MainDocumentPart mainDocPart = document.MainDocumentPart;

            FooterPart footerPart1 = mainDocPart.AddNewPart<FooterPart>("r98");



            Wp.Footer footer1 = new Wp.Footer();

            Wp.Paragraph paragraph1 = new Wp.Paragraph() { };



            Wp.Run run1 = new Wp.Run();
            Wp.Text text1 = new Wp.Text();
            text1.Text = footerText;

            Wp.FontSize wpFontSize = new Wp.FontSize { Val = (2 * fontSize).ToString() };  // Size in half points
            Wp.RunProperties runProperties = new Wp.RunProperties();
            Wp.RunFonts runFonts = new Wp.RunFonts { Ascii = "Calibri", HighAnsi = "Calibri" };
            runProperties.Append(runFonts);
            runProperties.Append(wpFontSize);
            run1.Append(runProperties);
            run1.Append(text1);

            Wp.ParagraphProperties paragraphProperties = new Wp.ParagraphProperties();
            Wp.Justification justification = new Wp.Justification() { Val = Wp.JustificationValues.Center };
            paragraphProperties.Append(justification);
            paragraph1.Append(paragraphProperties);

            paragraph1.Append(run1);


            footer1.Append(paragraph1);

            footerPart1.Footer = footer1;



            Wp.SectionProperties sectionProperties1 = mainDocPart.Document.Body.Descendants<Wp.SectionProperties>().FirstOrDefault();
            if (sectionProperties1 == null)
            {
                sectionProperties1 = new Wp.SectionProperties() { };

                Wp.PageMargin pageMargin = new Wp.PageMargin()
                {
                    Top = topMargin,
                    Right = rigthMargin,
                    Bottom = bottomMargin,
                    Left = leftMargin,
                    Header = 720U,
                    Footer = 640U,
                    Gutter = 0U
                };
                sectionProperties1.Append(pageMargin);


                mainDocPart.Document.Body.Append(sectionProperties1);
            }
            Wp.FooterReference footerReference1 = new Wp.FooterReference() { Type = DocumentFormat.OpenXml.Wordprocessing.HeaderFooterValues.Default, Id = "r98" };


            sectionProperties1.InsertAt(footerReference1, 0);

        }

        /// <summary>
        /// Creates a  Wp.Paragraph with the given lines, fontsizes and colors. Possibly bordered
        /// TODO: Font should be given as a paramter. Now Comic Sans MS is hardcoded
        /// TODO: Better handling of Font colors
        /// </summary>
        /// <param name="lines">Array with texts</param>
        /// <param name="fontSizes">Array with one font size for each line</param>
        /// <param name="colors">Array with one font color for each line</param>
        /// <param name="borders"></param>
        /// <returns></returns>

        public void test()
        {

        }
        /// <summary>
        /// Creates a  Wp.Paragraph with the given lines, fontsizes and colors. Possibly bordered
        /// TODO: Font should be given as a paramter. Now Comic Sans MS is hardcoded
        /// TODO: Better handling of Font colors
        /// TODO: bullet list css style, now hard coded, should be a parameter
        /// </summary>
        /// <param name="lines"></param>
        /// <param name="fontSizes"></param>
        /// <param name="colors"></param>
        /// <param name="pageBreakBefore"></param>
        /// <param name="borders"></param>
        /// <param name="distanceBefore"></param>
        /// <param name="distanceAfter"></param>
        /// <param name="underlineFirstRow"></param>
        /// <param name="bulletStart">The paragraph may have zero or one bullet lists </param>
        /// <param name="bulletLength"></param>
        /// <returns></returns>
        public static Wp.Paragraph GenerateParagraph(string[] lines, int[] fontSizes, String[] colors,
            Boolean pageBreakBefore, Wp.ParagraphBorders borders, int distanceBefore = 0, int distanceAfter = 0,
            Boolean underlineFirstRow = false)
        {
            //FlyerCreator.htmlStringBuilder.Append("<p class='m8'>\n");

            Wp.Paragraph paragraph = new Wp.Paragraph();
            Wp.ParagraphProperties paragraphProperties = new Wp.ParagraphProperties(borders);
            if (pageBreakBefore)
            {
                paragraphProperties.PageBreakBefore = new Wp.PageBreakBefore();
            }

            Wp.SpacingBetweenLines spacingBetweenLines = new Wp.SpacingBetweenLines();
            spacingBetweenLines.LineRule = Wp.LineSpacingRuleValues.Exact;
            spacingBetweenLines.Before = distanceBefore.ToString();
            spacingBetweenLines.After = distanceAfter.ToString();
            paragraphProperties.Append(spacingBetweenLines);


            Wp.Justification justification = new Wp.Justification() { Val = Wp.JustificationValues.Center };
            paragraphProperties.Append(justification);
            paragraph.Append(paragraphProperties);
            for (int i = 0; i < lines.Length; i++)
            {
                String line = lines[i];
                   String fontSizeTxt = (fontSizes[i] * 2).ToString();
                Wp.FontSize fontSize = new Wp.FontSize { Val = new OXML.StringValue(fontSizeTxt) };  // Size in half points

                Wp.Run run = new Wp.Run();
                Wp.RunProperties runProperties = new Wp.RunProperties();
                Wp.RunFonts runFonts = new Wp.RunFonts { Ascii = "Comic Sans MS", HighAnsi = "Comic Sans MS" };
                runProperties.Append(runFonts);
                runProperties.Append(fontSize);
                if (underlineFirstRow && i == 0)
                {
                    runProperties.Underline = new Wp.Underline() { Val = Wp.UnderlineValues.Single };
                }
                Wp.Text text1 = new Wp.Text();
                text1.Text = line;
                if (colors[i] == "Red")
                {
                    Wp.Color color = new Wp.Color() { Val = "FF0000" };
                    runProperties.Append(color);

                }
                run.Append(runProperties);
                run.Append(text1);
                if (i < lines.Length - 1)
                {
                    run.Append(new Wp.Break());
                }
                paragraph.Append(run);

            }
            //FlyerCreator.htmlStringBuilder.Append("</p>\n");

            return paragraph;
        }

        public static void CreateTableBorders(Wp.Table table, uint _size, Wp.BorderValues borderTypeOuter = Wp.BorderValues.Single,
            Wp.BorderValues borderTypeInner = Wp.BorderValues.None, String color = "E0C512")
        {
            OXML.UInt32Value size = (OXML.UInt32Value)_size;

            Wp.TableProperties props = new Wp.TableProperties(
                    new Wp.TableBorders(
                    new Wp.TopBorder
                    {
                        Val = new OXML.EnumValue<Wp.BorderValues>(borderTypeOuter),
                        Size = size,
                        Color = color
                    },
                    new Wp.BottomBorder
                    {
                        Val = new OXML.EnumValue<Wp.BorderValues>(borderTypeOuter),
                        Size = size,
                        Color = color
                    },
                    new Wp.LeftBorder
                    {
                        Val = new OXML.EnumValue<Wp.BorderValues>(borderTypeOuter),
                        Size = size,
                        Color = color
                    },
                    new Wp.RightBorder
                    {
                        Val = new OXML.EnumValue<Wp.BorderValues>(borderTypeOuter),
                        Size = size,
                        Color = color
                    },
                    new Wp.InsideHorizontalBorder
                    {
                        Val = new OXML.EnumValue<Wp.BorderValues>(borderTypeInner),
                        Size = size,
                        Color = color
                    },
                    new Wp.InsideVerticalBorder
                    {
                        Val = new OXML.EnumValue<Wp.BorderValues>(borderTypeInner),
                        Size = size,
                        Color = color
                    }));

            table.AppendChild<Wp.TableProperties>(props);
        }
        /// <summary>
        /// Sets table margins in Dxa: Twentieths of a Point.
        /// </summary>
        /// <param name="table"></param>
        /// <param name="topMargin"></param>
        /// <param name="startMargin"></param>
        /// <param name="bottomMargin"></param>
        /// <param name="endMargin"></param>
        public static void CreateTableMargins(Wp.Table table, int topMargin = 150, int startMargin = 50, int bottomMargin = 5, int endMargin = 50)
        {

            Wp.TableProperties tblProp = new Wp.TableProperties(
                // new Wp.TableCellSpacing() { Width = "200", Type = Wp.TableWidthUnitValues.Dxa },
                new Wp.TableCellMarginDefault(
                    new Wp.TopMargin() { Width = topMargin.ToString(), Type = Wp.TableWidthUnitValues.Dxa },
                    new Wp.StartMargin() { Width = startMargin.ToString(), Type = Wp.TableWidthUnitValues.Dxa },
                    new Wp.BottomMargin() { Width = bottomMargin.ToString(), Type = Wp.TableWidthUnitValues.Dxa },
                    new Wp.EndMargin() { Width = endMargin.ToString(), Type = Wp.TableWidthUnitValues.Dxa })
                    );
            table.AppendChild<Wp.TableProperties>(tblProp);
        }
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
            _hPO.Text = (x0_mm * mmToEmu).ToString();  // Convert mm to EMU
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

        public static void AddBulletList(List<string> sentences, MainDocumentPart mainPart, Wp.Body body, int fontSize)
        {
            var runList = ListOfStringToRunList(sentences, fontSize);

            AddBulletList(runList, mainPart, body);
        }
        private static List<Wp.Run> ListOfStringToRunList(List<string> sentences, int fontSize)
        {
            var runList = new List<Wp.Run>();
            foreach (string item in sentences)
            {
                var newRun = new Wp.Run();
                Wp.RunProperties runProperties = new Wp.RunProperties();
                Wp.FontSize wpFontSize = new Wp.FontSize { Val = (2 * fontSize).ToString() };  // Size in half points
                runProperties.Append(wpFontSize);
                Wp.RunFonts runFonts = new Wp.RunFonts { Ascii = "Comic Sans MS" };
                runProperties.Append(runFonts);
                newRun.Append(runProperties);
                newRun.AppendChild(new Wp.Text(item));
                runList.Add(newRun);
            }

            return runList;
        }


        public static void AddBulletList(List<Wp.Run> runList, MainDocumentPart mainPart, Wp.Body body)
        {
            //Introduce bulleted numbering in case it will be needed at some point
             NumberingDefinitionsPart numberingPart = mainPart.NumberingDefinitionsPart;
            if (numberingPart == null)
            {
                numberingPart = mainPart.AddNewPart<NumberingDefinitionsPart>("NumberingDefinitionsPart001");
                Wp.Numbering element = new Wp.Numbering();
                element.Save(numberingPart);
            }

            // Insert an AbstractNum into the numbering part numbering list.  The order seems to matter or it will not pass the 
            // Open XML SDK Productity Tools validation test.  AbstractNum comes first and then NumberingInstance and we want to
            // insert this AFTER the last AbstractNum and BEFORE the first NumberingInstance or we will get a validation error.
            var abstractNumberId = numberingPart.Numbering.Elements<Wp.AbstractNum>().Count() + 1;
            var abstractLevel = new Wp.Level(new Wp.NumberingFormat() { Val = Wp.NumberFormatValues.Bullet }, new Wp.LevelText() { Val = "·" }) { LevelIndex = 0 };
            var abstractNum1 = new Wp.AbstractNum(abstractLevel) { AbstractNumberId = abstractNumberId };

            if (abstractNumberId == 1)
            {
                numberingPart.Numbering.Append(abstractNum1);
            }
            else
            {
                Wp.AbstractNum lastAbstractNum = numberingPart.Numbering.Elements<Wp.AbstractNum>().Last();
                numberingPart.Numbering.InsertAfter(abstractNum1, lastAbstractNum);
            }

            // Insert an NumberingInstance into the numbering part numbering list.  The order seems to matter or it will not pass the 
            // Open XML SDK Productity Tools validation test.  AbstractNum comes first and then NumberingInstance and we want to
            // insert this AFTER the last NumberingInstance and AFTER all the AbstractNum entries or we will get a validation error.
            var numberId = numberingPart.Numbering.Elements<Wp.NumberingInstance>().Count() + 1;
            Wp.NumberingInstance numberingInstance1 = new Wp.NumberingInstance() { NumberID = numberId };
            Wp.AbstractNumId abstractNumId1 = new Wp.AbstractNumId() { Val = abstractNumberId };
            numberingInstance1.Append(abstractNumId1);

            if (numberId == 1)
            {
                numberingPart.Numbering.Append(numberingInstance1);
            }
            else
            {
                var lastNumberingInstance = numberingPart.Numbering.Elements<Wp.NumberingInstance>().Last();
                numberingPart.Numbering.InsertAfter(numberingInstance1, lastNumberingInstance);
            }

            

            foreach (Wp.Run runItem in runList)
            {
                // Create items for paragraph properties
                var numberingProperties = new Wp.NumberingProperties(new Wp.NumberingLevelReference() { Val = 0 }, new Wp.NumberingId() { Val = numberId });
                var spacingBetweenLines1 = new Wp.SpacingBetweenLines() { After = "0" };  // Get rid of space between bullets
                var indentation = new Wp.Indentation() { Left = "720", Hanging = "360" };  // correct indentation 

                Wp.ParagraphMarkRunProperties paragraphMarkRunProperties1 = new Wp.ParagraphMarkRunProperties();
                Wp.RunFonts runFonts1 = new Wp.RunFonts() { Ascii = "Symbol", HighAnsi = "Symbol" };
                paragraphMarkRunProperties1.Append(runFonts1);

                // create paragraph properties
                var paragraphProperties = new Wp.ParagraphProperties(numberingProperties, spacingBetweenLines1, indentation, paragraphMarkRunProperties1);

                // Create paragraph 
                var newPara = new Wp.Paragraph(paragraphProperties);

                // Add run to the paragraph
                newPara.AppendChild(runItem);
                // Add one bullet item to the body
               body.AppendChild(newPara);
            }
        }

    }
}
