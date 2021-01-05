using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.Content;
using PdfSharp.Pdf.IO;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace TestConsoleApplication
{
    /// <summary>
    /// To Install This Package <code>Install-Package PDFsharp-MigraDoc -Version 1.50.4845-RC2a</code>
    /// Tutorial <c>http://www.pdfsharp.net/wiki/PDFsharpSamples.ashx</c>
    /// </summary>
    class PDF_Sharp_Migra_Doc
    {
        /// <summary>
        /// Tutorial Link http://www.pdfsharp.net/wiki/HelloWorld-sample.ashx
        /// </summary>
        public static void HelloWorld()
        {
            // Create a new PDF document
            PdfDocument document = new PdfDocument();
            document.Info.Title = "Created with PDFsharp";

            // Create an empty page
            PdfPage page = document.AddPage();

            // Get an XGraphics object for drawing
            XGraphics gfx = XGraphics.FromPdfPage(page);

            // Create a font
            XFont font = new XFont("Verdana", 48, XFontStyle.BoldItalic);

            // Draw the text
            gfx.DrawString("Hello, World!", font, XBrushes.Plum, new XRect(0, 0, page.Width, page.Height), XStringFormats.Center);

            //Draw Square
            gfx.DrawLines(XPens.DarkRed, new XPoint[] { new XPoint(200, 500), new XPoint(300, 500), new XPoint(300, 600), new XPoint(200, 600), new XPoint(200, 500) });
            gfx.DrawLines(XPens.DarkRed, new XPoint[] { new XPoint(250, 479.5), new XPoint(320.5, 550), new XPoint(250, 620.5), new XPoint(179.5, 550), new XPoint(250, 479.5) });
            gfx.DrawArc(XPens.DarkRed, new XRect(180, 480, 140, 140), 90, 0);
            //gfx.DrawArc(XPens.DarkCyan, new XRect(180, 485, 145, 145), 90, 0);

            // Save the document...
            const string filename = @"..\..\..\HelloWorld.pdf";
            document.Save(filename);

            // ...and start a viewer.
            Process.Start(filename);
        }        
    }
}
