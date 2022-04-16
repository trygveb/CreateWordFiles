using System;
using System.IO;

using OXML= DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Wp= DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace CreateWordFiles
{
    internal class Creator
    {
        public static void CreateWordprocessingDocument(string filepath, String danceName, String danceDates)
        {
            // Create a document by supplying the filepath. 
            using (WordprocessingDocument wordDocument =
                WordprocessingDocument.Create(filepath, OXML.WordprocessingDocumentType.Document))
            {
                // Add a main document part. 
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();

                // Create the document structure and add some text.
                mainPart.Document = new Wp.Document();
                Wp.Body body = mainPart.Document.AppendChild(new Wp.Body());

                Wp.Paragraph paragraph1 = GenerateParagraph1(danceName.ToUpper(), danceDates);
                body.AppendChild(paragraph1);

                String fileNameLogo = @"D:\Mina dokument\Sqd\Motiv8s\Badge\M8-logo1.gif";
                addImage("Anchor", wordDocument, fileNameLogo, 0.1667, 1.0, 1.6);

                String fileNameCaller = @"D:\Mina dokument\Sqd\Motiv8s\Dokument\Flyers\jesper.gif";
                addImage("Inline", wordDocument, fileNameCaller, 0.8, 6.0, 10.0);

                body.AppendChild(GenerateParagraph2());
                body.AppendChild(GenerateParagraph4());
                body.AppendChild(GenerateParagraph5());
                body.AppendChild(GenerateParagraph6());
            }
        }
        public static Wp.Paragraph GenerateParagraph1(String danceName, String danceDates)
        {
            //Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "004F7104", RsidParagraphProperties = "008F2986", RsidRunAdditionDefault = "004F7104" };
            Wp.Paragraph paragraph1 = new Wp.Paragraph();


            String[] lines= { "VÄLKOMNA", "till", danceName, danceDates };
            int[] fontSizes = { 20, 12, 32, 20 };
            return GenerateParagraph(lines, fontSizes);
        }
        public static Wp.Paragraph GenerateParagraph2()
        {
            String[] lines = { "Jesper Wilhelmsson" };
            int[] fontSizes = { 32 };
            return GenerateParagraph(lines, fontSizes);
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
            return GenerateParagraph(lines, fontSizes);
        }
        public static Wp.Paragraph GenerateParagraph5()
        {
            String[] lines = {
                "Plats: Segersjö Folkets Hus, Scheelevägen 41 i Tumba"
            };
            int[] fontSizes = { 14 };
            return GenerateParagraph(lines, fontSizes);
        }
        public static Wp.Paragraph GenerateParagraph6()
        {
            String[] lines = {
                "Vi använder rotationsprogram på samtliga nivåer!",
                "Ta med eget fika! - Vi ordnar hämtning av Pizza och sallad till pausen!"
            };
            int[] fontSizes = { 14, 14 };
            return GenerateParagraph(lines, fontSizes);
        }

        private static Wp.Paragraph GenerateParagraph(string[] lines, int[] fontSizes)
        {
            Wp.Paragraph paragraph = new Wp.Paragraph();
            for (int i = 0; i < lines.Length; i++)
            {
                String fontSizeTxt= (fontSizes[i]*2).ToString();
                String line = lines[i];
                Wp.FontSize fontSize = new Wp.FontSize { Val = new OXML.StringValue(fontSizeTxt) };  // Size in half points
                Wp.ParagraphProperties paragraphProperties = new Wp.ParagraphProperties();
                Wp.Justification justification = new Wp.Justification() { Val = Wp.JustificationValues.Center };
                paragraphProperties.Append(justification);

                Wp.Run run = new Wp.Run();
                Wp.RunProperties runProperties = new Wp.RunProperties();
                Wp.RunFonts runFonts = new Wp.RunFonts { Ascii = "Comic Sans MS" };
                runProperties.Append(runFonts);
                runProperties.Append(fontSize);
                Wp.Text text1 = new Wp.Text();
                text1.Text = line;
                run.Append(runProperties);
                run.Append(text1);
                run.Append(new Wp.Break());
                paragraph.Append(paragraphProperties);
                paragraph.Append(run);

            }
            return paragraph;
        }

        public static void addImage(String type, WordprocessingDocument wordprocessingDocument, String fileName, double scale, double x0_cm, double y0_cm)
        {
            MainDocumentPart mainPart = wordprocessingDocument.MainDocumentPart;
            ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);
            //String fileName = @"D:\Mina dokument\Sqd\Motiv8s\Badge\M8-logo1.gif";
            int iWidth = 0;
            int iHeight = 0;
            using (System.Drawing.Bitmap bmp = new System.Drawing.Bitmap(fileName))
            {
                iWidth = bmp.Width;
                iHeight = bmp.Height;
            }
            using (FileStream stream = new FileStream(fileName, FileMode.Open))
            {
                imagePart.FeedData(stream);
            }
            if (type == "Anchor")
            {
                AddAnchorImageToBody(wordprocessingDocument, mainPart.GetIdOfPart(imagePart), (int)(scale * iWidth), (int)(scale * iHeight), x0_cm, y0_cm);
            } else
            {
                AddInlineImageToBody(wordprocessingDocument, mainPart.GetIdOfPart(imagePart), (int)(scale * iWidth), (int)(scale * iHeight));
            }
        }
        public static void AddAnchorImageToBody(WordprocessingDocument wordprocessingDocument, String relationshipId, int iWidth, int iHeight, double x0_cm, double y0_cm)
        {
             // Define the reference of the image.
            var x = new DW.Anchor();
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
            long x0_emu = (long)Math.Round(x0_cm * f1);
            long y0_emu = (long)Math.Round(y0_cm * f1);
            // convert the pixels to EMUs this way:
            long iWidth = (long)Math.Round((decimal)wPixels * 9525);
            long iHeight = (long)Math.Round((decimal)hPixels * 9525);

            Wp.Drawing _drawing = new Wp.Drawing();
            DW.Anchor _anchor = new DW.Anchor()
            {
                DistanceFromTop = (OXML.UInt32Value)0U,
                DistanceFromBottom = (OXML.UInt32Value)0U,
                DistanceFromLeft = (OXML.UInt32Value)0U,
                DistanceFromRight = (OXML.UInt32Value)0U,
                SimplePos = true,
                RelativeHeight = (OXML.UInt32Value)0U,
                BehindDoc = false,
                Locked = false,
                LayoutInCell = true,
                AllowOverlap = true,
                EditId = "44CEF5E4",
                AnchorId = "44803ED1"
            };
            DW.SimplePosition _spos = new DW.SimplePosition()
            {
                X = x0_emu,
                Y = y0_emu  
            };

            DW.HorizontalPosition _hp = new DW.HorizontalPosition()
            {
                RelativeFrom = DW.HorizontalRelativePositionValues.Column
            };
            DW.PositionOffset _hPO = new DW.PositionOffset();
            // _hPO.Text = "76200";
            _hPO.Text = "0";
            _hp.Append(_hPO);

            DW.VerticalPosition _vp = new DW.VerticalPosition()
            {
                RelativeFrom = DW.VerticalRelativePositionValues.Paragraph
            };
            DW.PositionOffset _vPO = new DW.PositionOffset();
            //_vPO.Text = "83820";
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

            DW.WrapTight _wp = new DW.WrapTight()
            {
                WrapText = DW.WrapTextValues.BothSides
            };

            DW.WrapPolygon _wpp = new DW.WrapPolygon()
            {
                Edited = false
            };
            DW.StartPoint _sp = new DW.StartPoint()
            {
                X = 0L,
                Y = 0L
            };

            DW.LineTo _l1 = new DW.LineTo() { X = 0L, Y = 0L };
            DW.LineTo _l2 = new DW.LineTo() { X = 0L, Y = 0L };
            DW.LineTo _l3 = new DW.LineTo() { X = 0L, Y = 0L };
            DW.LineTo _l4 = new DW.LineTo() { X = 0L, Y = 0L };

            _wpp.Append(_sp);
            _wpp.Append(_l1);
            _wpp.Append(_l2);
            _wpp.Append(_l3);
            _wpp.Append(_l4);

            _wp.Append(_wpp);

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
            _anchor.Append(_wp);
            _anchor.Append(_dp);
            _anchor.Append(_g);

            _drawing.Append(_anchor);

            return _drawing;
        }
    }
}
