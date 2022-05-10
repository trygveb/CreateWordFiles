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
        /// <summary>
        /// Adds a footer to a document
        /// </summary>
        /// <param name="document"></param>
        /// <param name="footerText"></param>
        /// <param name="fontSize">Font size in points</param>
        public static void ApplyFooter(WordprocessingDocument document, String footerText, int fontSize)
        {
            // Get the main document part.
            MainDocumentPart mainDocPart = document.MainDocumentPart;

            FooterPart footerPart1 = mainDocPart.AddNewPart<FooterPart>("r98");



            Wp.Footer footer1 = new Wp.Footer();

            Wp.Paragraph paragraph1 = new Wp.Paragraph() { };



            Wp.Run run1 = new Wp.Run();
            Wp.Text text1 = new Wp.Text();
            text1.Text = footerText;

            Wp.FontSize wpFontSize = new Wp.FontSize { Val = (2*fontSize).ToString() };  // Size in half points
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

                Wp.PageMargin pageMargin = new Wp.PageMargin() { Top = 200, Right =1008U, Bottom = 1008, Left = 1008U, Header = 720U,
                    Footer = 640U, Gutter = 0U };
                sectionProperties1.Append(pageMargin);


                mainDocPart.Document.Body.Append(sectionProperties1);
            }
            Wp.FooterReference footerReference1 = new Wp.FooterReference() { Type = DocumentFormat.OpenXml.Wordprocessing.HeaderFooterValues.Default, Id = "r98" };


            sectionProperties1.InsertAt(footerReference1, 0);

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
    }
}
