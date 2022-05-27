using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using Pic = DocumentFormat.OpenXml.Drawing.Pictures;

namespace CreateWordFiles
{
    public class AddCallerPicture
    {
        // Creates an Paragraph instance and adds its children.
        public static Paragraph GenerateParagraph()
        {
            Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "00C91193", RsidRunAdditionDefault = "00B10F21" };

            ParagraphProperties paragraphProperties1 = new ParagraphProperties();

            ParagraphBorders paragraphBorders1 = new ParagraphBorders();
            TopBorder topBorder1 = new TopBorder() { Val = BorderValues.None, Color = "E0C512", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            LeftBorder leftBorder1 = new LeftBorder() { Val = BorderValues.None, Color = "E0C512", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder1 = new BottomBorder() { Val = BorderValues.None, Color = "E0C512", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            RightBorder rightBorder1 = new RightBorder() { Val = BorderValues.None, Color = "E0C512", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            paragraphBorders1.Append(topBorder1);
            paragraphBorders1.Append(leftBorder1);
            paragraphBorders1.Append(bottomBorder1);
            paragraphBorders1.Append(rightBorder1);
            Justification justification1 = new Justification() { Val = JustificationValues.Center };

            paragraphProperties1.Append(paragraphBorders1);
            paragraphProperties1.Append(justification1);

            Run run1 = new Run();

            RunProperties runProperties1 = new RunProperties();
            NoProof noProof1 = new NoProof();

            runProperties1.Append(noProof1);

            Drawing drawing1 = new Drawing();

            Wp.Inline inline1 = new Wp.Inline() { DistanceFromTop = (UInt32Value)1000U, DistanceFromBottom = (UInt32Value)0U, DistanceFromLeft = (UInt32Value)0U, DistanceFromRight = (UInt32Value)0U, EditId = "50D07946" };
            Wp.Extent extent1 = new Wp.Extent() { Cx = 2552700L, Cy = 2371725L };
            Wp.EffectExtent effectExtent1 = new Wp.EffectExtent() { LeftEdge = 38100L, TopEdge = 38100L, RightEdge = 19050L, BottomEdge = 28575L };
            Wp.DocProperties docProperties1 = new Wp.DocProperties() { Id = (UInt32Value)2U, Name = "Picture 1" };

            Wp.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties1 = new Wp.NonVisualGraphicFrameDrawingProperties();

            A.GraphicFrameLocks graphicFrameLocks1 = new A.GraphicFrameLocks() { NoChangeAspect = true };
            graphicFrameLocks1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            nonVisualGraphicFrameDrawingProperties1.Append(graphicFrameLocks1);

            A.Graphic graphic1 = new A.Graphic();
            graphic1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            A.GraphicData graphicData1 = new A.GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" };
            
            Pic.Picture picture1 = new Pic.Picture();
            picture1.AddNamespaceDeclaration("pic", "http://schemas.openxmlformats.org/drawingml/2006/picture");

            Pic.NonVisualPictureProperties nonVisualPictureProperties1 = new Pic.NonVisualPictureProperties();
            Pic.NonVisualDrawingProperties nonVisualDrawingProperties1 = new Pic.NonVisualDrawingProperties() { Id = (UInt32Value)0U, Name = "https://www.motiv8s.se/19/images/callers/david_heffron_USA.jpg" };
            Pic.NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties1 = new Pic.NonVisualPictureDrawingProperties();

            nonVisualPictureProperties1.Append(nonVisualDrawingProperties1);
            nonVisualPictureProperties1.Append(nonVisualPictureDrawingProperties1);

            Pic.BlipFill blipFill1 = new Pic.BlipFill();

            A.Blip blip1 = new A.Blip() { Embed = "rId8", CompressionState = A.BlipCompressionValues.Print };
            A.BlipExtensionList blipExtensionList1 = new A.BlipExtensionList();

            blip1.Append(blipExtensionList1);

            A.Stretch stretch1 = new A.Stretch();
            A.FillRectangle fillRectangle1 = new A.FillRectangle();

            stretch1.Append(fillRectangle1);

            blipFill1.Append(blip1);
            blipFill1.Append(stretch1);

            Pic.ShapeProperties shapeProperties1 = new Pic.ShapeProperties();

            A.Transform2D transform2D1 = new A.Transform2D();
            A.Offset offset1 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents1 = new A.Extents() { Cx = 2552700L, Cy = 2371725L };

            transform2D1.Append(offset1);
            transform2D1.Append(extents1);

            A.PresetGeometry presetGeometry1 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList1 = new A.AdjustValueList();

            presetGeometry1.Append(adjustValueList1);

            A.Outline outline1 = new A.Outline() { Width = 38100 };

            A.SolidFill solidFill1 = new A.SolidFill();
            A.SchemeColor schemeColor1 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent4 };

            solidFill1.Append(schemeColor1);

            outline1.Append(solidFill1);

            shapeProperties1.Append(transform2D1);
            shapeProperties1.Append(presetGeometry1);
            shapeProperties1.Append(outline1);

            picture1.Append(nonVisualPictureProperties1);
            picture1.Append(blipFill1);
            picture1.Append(shapeProperties1);

            graphicData1.Append(picture1);

            graphic1.Append(graphicData1);

            inline1.Append(extent1);
            inline1.Append(effectExtent1);
            inline1.Append(docProperties1);
            inline1.Append(nonVisualGraphicFrameDrawingProperties1);
            inline1.Append(graphic1);

            drawing1.Append(inline1);

            run1.Append(runProperties1);
            run1.Append(drawing1);
            BookmarkStart bookmarkStart1 = new BookmarkStart() { Name = "_GoBack", Id = "0" };
            BookmarkEnd bookmarkEnd1 = new BookmarkEnd() { Id = "0" };

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(run1);
            paragraph1.Append(bookmarkStart1);
            paragraph1.Append(bookmarkEnd1);
            return paragraph1;
        }


    }
}
