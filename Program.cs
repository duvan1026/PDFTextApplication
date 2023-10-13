using System;
using System.Diagnostics;
using System.Drawing;
using PdfSharp.Drawing;
using PdfSharp.Drawing.Layout;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using Tesseract;

namespace PDFTextApplication
{
    internal class Program
    {

        static void Main()
        {

            //string tiffImagePath = "C:\\Users\\duvan.castro\\Desktop\\TestPDFText\\InputFile\\PASIVO00001086F.tif";
            string tiffImagePath = "C:\\Users\\duvan.castro\\Desktop\\TestPDFText\\InputFile\\PBOGESCANER01TripleA100423214\\Imagen6.tif";
            string pdfOutputPath = "C:\\Users\\duvan.castro\\Desktop\\TestPDFText\\OutputFile\\OutImagen640.pdf";

            // Ruta al archivo PDF original
            string pdfPath = "C:\\Users\\duvan.castro\\Desktop\\TestPDFText\\InputFile\\Imagen642.pdf";

            // Procesar el PDF
            ConvertTiffToPdf(tiffImagePath, pdfOutputPath);

            Console.WriteLine("Proceso completado. El PDF de texto se ha guardado.");
        }
        static void ConvertTiffToPdf(string tiffImagePath, string pdfOutputPath)
        {
            string environmentVariable = "TESSDATA_PREFIX";
            string tessdataPath = @"C:\Program Files (x86)\Tesseract-OCR\tessdata";
            string language = "spa";

            // Establecemos una variable de entorno(NombreVariableEntorno, ValorAsignado)
            Environment.SetEnvironmentVariable(environmentVariable, tessdataPath);

            using (var engine = new TesseractEngine(tessdataPath, language, EngineMode.Default))
            {
                using (var image = new Bitmap(tiffImagePath))
                {
                    using (var pageProcessor = engine.Process(image))
                    {
                        var iter = pageProcessor.GetIterator();
                        iter.Begin();

                        double imageWidth = image.Width;                     // Ancho de la imagen 
                        double imageHeight = image.Height;                   // Alto de la imagen
                        
                        var document = new PdfDocument();                    // Crear un nuevo documento PDF
                        var page = document.AddPage();                       // Se crea una nueva pagina

                        double dpi = 96.0;                                  // Resolución estándar de pantalla

                        double widthInPoints = imageWidth * 72.0 / dpi;
                        double heightInPoints = imageHeight * 72.0 / dpi;

                        page.Width = widthInPoints;
                        page.Height = heightInPoints;
                        var gfx = XGraphics.FromPdfPage(page);

                        double scaleWidth = widthInPoints / image.Width;
                        double scaleHeight = heightInPoints / image.Height;

                        // Agregar la imagen TIFF al PDF
                        var xImage = XImage.FromFile(tiffImagePath);
                        gfx.DrawImage(xImage, 0, 0, widthInPoints, heightInPoints);


                        while (iter.Next(PageIteratorLevel.Word))
                        {
                            var word = iter.GetText(PageIteratorLevel.Word).Trim(); // Eliminar espacios en blanco al principio y al final


                            if (!string.IsNullOrEmpty(word) && !word.Contains("|"))
                            {
                                Rect bounds;
                                if (iter.TryGetBoundingBox(PageIteratorLevel.Word, out bounds))
                                {
                                    // Obtener las coordenadas del carácter
                                    double x1 = bounds.X1 * scaleWidth;
                                    double y1 = (bounds.Y1 * scaleHeight) + (bounds.Height * scaleHeight);


                                    double width = bounds.Width * scaleWidth;
                                    double height = bounds.Height * scaleHeight;

                                    // Calcular el tamaño de fuente con el coeficiente de escala
                                    double realSizeInPoints = 12;


                                    string fontFamilyName = "Arial";                             // Nombre de la fuente
                                    double targetHeightInPixels = bounds.Height;                            // Altura en píxeles que deseas alcanzar
                                    double epsilon = 1;                                          // Margen de error permitido

                                    var fontTest = new XFont(fontFamilyName, realSizeInPoints);
                                    double fontHeightInPixels = fontTest.GetHeight();            // Obtener la altura en píxeles de la fuente

                                    while (Math.Abs(fontHeightInPixels - targetHeightInPixels) > epsilon)
                                    {
                                        if (fontHeightInPixels < targetHeightInPixels)
                                        {                                            
                                            realSizeInPoints += 1.0;   // Incrementa el tamaño de fuente
                                        }
                                        else
                                        {                                            
                                            realSizeInPoints -= 1.0;  // Decrementa el tamaño de fuente
                                        }

                                        // Vuelve a calcular la altura en píxeles con el nuevo tamaño de fuente
                                        fontTest = new XFont(fontFamilyName, realSizeInPoints);
                                        fontHeightInPixels = fontTest.GetHeight();
                                    }



                                    //realSizeInPoints = realSizeInPoints * 



                                    //var fontTest = new XFont(fontFamilyName, realSizeInPoints);
                                    //double fontHeightInPixels = fontTest.GetHeight();            // Obtener la altura en píxeles de la fuente


                                    //double factorScaler = (realSizeInPoints * height) / fontHeightInPixels;
                                    //Console.WriteLine("-----------------");
                                    //Console.WriteLine($"La altura de la fuente {fontFamilyName} a {realSizeInPoints} puntos es {fontHeightInPixels} píxeles.");
                                    //realSizeInPoints = factorScaler;
                                    //fontTest = new XFont(fontFamilyName, realSizeInPoints);
                                    //fontHeightInPixels = fontTest.GetHeight();  // Obtener la altura en píxeles de la fuente
                                    Console.WriteLine($"La altura de la fuente {fontFamilyName} a {realSizeInPoints} puntos es {fontHeightInPixels} píxeles.");



                                    if (bounds.X1 < imageWidth && bounds.Y1 < imageHeight)
                                    {
                                        var font = new XFont("Arial", realSizeInPoints);                 // Agrega la palabra con el tamaño de fuente calculado
                                        XBrush brush = XBrushes.Transparent;  // Establecer el color del texto como transparente
                                        gfx.DrawString(word, font, brush, new XPoint(x1, y1));  // Agregar la palabra y sus coordenadas al PDF
                                        //gfx.DrawString(word, font, XBrushes.Black, new XPoint(x1, y1));  // Agregar la palabra y sus coordenadas al PDF

                                    }
                                    else
                                    {
                                        Console.Write("Ha soprepasado los limites de la imagen ----------");
                                    }

                                    Console.Write($"Palabra: {word}");
                                    Console.Write($"  Tamaño Fuente: {realSizeInPoints}");
                                    Console.WriteLine($":  Ubicación: X={bounds.X1}, Y={bounds.Y1}, Ancho={bounds.Width}, Alto={bounds.Height}");
                                }

                            }

                        }

                        // Guardar el documento PDF
                        document.Save(pdfOutputPath);

                        Console.ReadLine();
                    }
                }
            }
        }
    }
}