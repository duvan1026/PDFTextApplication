using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using PdfSharp.Drawing;
using PdfSharp.Drawing.Layout;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using Tesseract;

namespace PDFTextApplication
{
    internal class Program
    {
        #region Constantes

        const string environmentVariable = "TESSDATA_PREFIX";
        const string tessdataPath = @"C:\Program Files (x86)\Tesseract-OCR\tessdata";
        const string language = "spa";
        const string _tiffImagePath = "C:\\Users\\duvan.castro\\Desktop\\TestPDFText\\InputFile\\PBOGESCANER01TripleA100423214\\Imagen6.tif";
        const string pdfOutputPath = "C:\\Users\\duvan.castro\\Desktop\\TestPDFText\\OutputFile\\OutImagen640.pdf";
        const double dpi = 96.0;                                  // Resolución estándar de pantalla
        
        #endregion

        static void Main()
        {
            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start(); // Comienza a medir el tiempo

            string tiffFolder = @"C:\Users\duvan.castro\Desktop\TestPDFText\InputFile\PBOGESCANER01TripleA100423214"; // Reemplaza con la ruta de tu carpeta

            string[] tiffFiles = Directory.GetFiles(tiffFolder, "*.tif");

            foreach (string tiffFile in tiffFiles)
            {
                string outputPath = Path.Combine(tiffFolder, Path.GetFileNameWithoutExtension(tiffFile) + ".Procesado.pdf");
                ConvertTiffToPdf(tiffFile, outputPath);                // Procesar el PDF
            }

            stopwatch.Stop(); // Detiene la medición
            TimeSpan elapsedTime = stopwatch.Elapsed;

            Console.WriteLine("Proceso completado en " + elapsedTime.TotalMilliseconds + " milisegundos.");
            Console.WriteLine("Proceso completado. El PDF de texto se ha guardado.");
            Console.ReadLine();
        }
        static void ConvertTiffToPdf(string tiffImagePath, string pdfOutputPath)
        {
            Environment.SetEnvironmentVariable(environmentVariable, tessdataPath);

            using (var engine = new TesseractEngine(tessdataPath, language, EngineMode.Default))
            {
                using (var image = new Bitmap(tiffImagePath))
                {
                    using (var pageProcessor = engine.Process(image))
                    {
                        double scaleWidth, scaleHeight;

                        double widthInPoints = ConvertToPoints(image.Width, dpi);
                        double heightInPoints = ConvertToPoints(image.Height, dpi);

                        var document = CreatePdfDocument(widthInPoints, heightInPoints);
                        var gfx = CreateGraphics(document);

                        scaleWidth = CalculateScale(widthInPoints, image.Width);
                        scaleHeight = CalculateScale(heightInPoints, image.Height);

                        AddImageToPdf(gfx, tiffImagePath, widthInPoints, heightInPoints);
                        ProcessText(pageProcessor, gfx, scaleWidth, scaleHeight);
                        SaveAndClosePdfDocument(document, pdfOutputPath);
                    }
                }
            }
        }

        static void SaveAndClosePdfDocument(PdfDocument document, string outputPath)
        {
            document.Save(outputPath);
            document.Close();
        }

        static XGraphics CreateGraphics(PdfDocument document)
        {
            var page = document.Pages[0];
            return XGraphics.FromPdfPage(page);
        }

        static double CalculateScale(double value1, double value2)
        {
            return value1 / value2;
        }

        static PdfDocument CreatePdfDocument(double width, double height)
        {
            var document = new PdfDocument();
            var page = document.AddPage();
            page.Width = width;
            page.Height = height;
            return document;
        }

        static double ConvertToPoints(double value, double dpi)
        {
            return value * 72.0 / dpi;
        }

        static void ProcessText(Tesseract.Page pageProcessor, XGraphics gfx, double scaleWidth, double scaleHeight)
        {
            var iter = pageProcessor.GetIterator();
            iter.Begin();

            while (iter.Next(PageIteratorLevel.Word))
            {
                var word = iter.GetText(PageIteratorLevel.Word).Trim();
                if (!string.IsNullOrEmpty(word) && !word.Contains("|"))
                {
                    Rect bounds;
                    if (iter.TryGetBoundingBox(PageIteratorLevel.Word, out bounds))
                    {
                        // Obtener las coordenadas del carácter
                        double x1 = bounds.X1 * scaleWidth;
                        double y1 = (bounds.Y1 * scaleHeight) + (bounds.Height * scaleHeight);

                        double realSizeInPoints = CalculateFontSize(iter, bounds.Height);

                        var font = new XFont("Arial", realSizeInPoints);                   // Agrega la palabra con el tamaño de fuente calculado
                        XBrush brush = XBrushes.Transparent;                               // Establecer el color del texto como transparente
                        gfx.DrawString(word, font, brush, new XPoint(x1, y1));             // Agregar la palabra y sus coordenadas al PDF
                                                                                           //gfx.DrawString(word, font, XBrushes.Black, new XPoint(x1, y1));  // Agregar la palabra y sus coordenadas al PDF

                        //Console.Write($"Palabra: {word}");
                        //Console.Write($"  Tamaño Fuente: {realSizeInPoints}");
                        //Console.WriteLine($":  Ubicación: X={bounds.X1}, Y={bounds.Y1}, Ancho={bounds.Width}, Alto={bounds.Height}");
                    }
                }
            }
        }

        static void AddImageToPdf(XGraphics gfx, string imagePath, double width, double height)
        {
            var xImage = XImage.FromFile(imagePath);
            gfx.DrawImage(xImage, 0, 0, width, height);
        }

        static double CalculateFontSize(PageIterator iter, double targetHeightInPixels)
        {
            const double epsilon = 1;
            const string fontFamilyName = "Arial";
            double realSizeInPoints = 12;

            XFont fontTest = new XFont(fontFamilyName, realSizeInPoints);
            double fontHeightInPixels = fontTest.GetHeight();

            while (Math.Abs(fontHeightInPixels - targetHeightInPixels) > epsilon)
            {
                if (fontHeightInPixels < targetHeightInPixels)
                {
                    realSizeInPoints += 1.0;
                }
                else
                {
                    realSizeInPoints -= 1.0;
                }

                fontTest = new XFont(fontFamilyName, realSizeInPoints);
                fontHeightInPixels = fontTest.GetHeight();
            }

            return realSizeInPoints;
        }
    }
}