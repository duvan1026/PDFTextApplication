using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Security.AccessControl;
using System.Security.Principal;
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

        const string tessdataPath = @"C:\Program Files (x86)\Tesseract-OCR\tessdata";
        const string language = "spa";
        const double dpi = 96.0;                                  // Resolución estándar de pantalla

        const string inputFile = @"C:\Users\duvan.castro\Desktop\TestPDFText\Data"; // Reemplaza con la ruta de tu carpeta
        const string outputFile = @"C:\Users\duvan.castro\Desktop\TestPDFText\Data";
        const string nameDirectoryDestination = @"Data.Process";

        #endregion

        static void Main()
        {
            // Iniciar el cronómetro para medir el tiempo de ejecución
            Stopwatch stopwatch = Stopwatch.StartNew();

            // Crear el directorio de destino para el archivo de salida
            string newOutputFile = Path.Combine(outputFile, nameDirectoryDestination);      
            CreateDirectoryWithWriteAccess(newOutputFile);

            // Generar un nombre de directorio basado en la fecha y hora actual
            string newOutputDirectoryDate = Path.Combine(newOutputFile, GetNameFileDate());
            CreateDirectoryWithWriteAccess(newOutputDirectoryDate);

            // Recorrer los directorios de entrada
            foreach (string currentDirectory in Directory.GetDirectories(inputFile))
            {
                string nameCurrentDirectory = Path.GetFileName(currentDirectory);

                // Verificar si el directorio actual no es el directorio de destino
                if (nameCurrentDirectory != nameDirectoryDestination)
                {
                    // Crear la ruta de destino para el directorio actual
                    string newRouteDestination = Path.Combine(newOutputDirectoryDate, nameCurrentDirectory);
                    CreateDirectoryWithWriteAccess(newRouteDestination);

                    // Obtener archivos TIFF en el directorio actual
                    string[] tiffFiles = Directory.GetFiles(currentDirectory, "*.tif");

                    // Procesar cada archivo TIFF y convertirlo a PDF
                    foreach (string tiffFile in tiffFiles)
                    {
                        string outputPath = Path.Combine(newRouteDestination, Path.GetFileNameWithoutExtension(tiffFile) + ".Procesado.pdf");
                        ConvertTiffToPdf(tiffFile, outputPath);                
                    }
                }
            }

            // Detener el cronómetro y calcular el tiempo transcurrido
            stopwatch.Stop(); 
            TimeSpan elapsedTime = stopwatch.Elapsed;
            Console.WriteLine($"Proceso completado en {elapsedTime.TotalMinutes} minutos {elapsedTime.Seconds} segundos");
            Console.WriteLine("Proceso completado. El PDF de texto se ha guardado.");
            Console.ReadLine();
        }


        static void ConvertTiffToPdf(string tiffImagePath, string pdfOutputPath)
        {
            Environment.SetEnvironmentVariable("TESSDATA_PREFIX", tessdataPath);

            using (var engine = new TesseractEngine(tessdataPath, language, EngineMode.Default))
            {
                using (var image = new Bitmap(tiffImagePath))
                {
                    using (var pageProcessor = engine.Process(image))
                    {
                        double scaleWidth, scaleHeight;

                        double widthInPoints = ConvertToPoints(image.Width, dpi);
                        double heightInPoints = ConvertToPoints(image.Height, dpi);

                        using (var document = CreatePdfDocument(widthInPoints, heightInPoints))
                        {
                            using (var gfx = CreateGraphics(document))
                            {
                                scaleWidth = CalculateScale(widthInPoints, image.Width);
                                scaleHeight = CalculateScale(heightInPoints, image.Height);

                                AddImageToPdf(gfx, tiffImagePath, widthInPoints, heightInPoints);
                                ProcessText(pageProcessor, gfx, scaleWidth, scaleHeight);
                                SaveAndClosePdfDocument(document, pdfOutputPath);
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Verifica si un directorio existe y lo crea si no existe. Luego, asigna permisos de escritura al directorio.
        /// </summary>
        /// <param name="directoryPath">La ruta del directorio a verificar y crear si es necesario.</param>
        public static void CreateDirectoryWithWriteAccess(string directoryPath)
        {
            if (!Directory.Exists(directoryPath))
            {
                Directory.CreateDirectory(directoryPath);
                AssignWritePrivilegesToDirectory(directoryPath);
            }
        }

        /// <summary>
        /// Genera un nombre de carpeta único basado en la fecha y la hora actual.
        /// </summary>
        /// <returns>Un nombre de carpeta en formato "yyyyMMddHHmmssfff".</returns>
        static string GetNameFileDate()
        {
            DateTime now = DateTime.Now;
            string nombreCarpeta = now.ToString("yyyyMMddHHmmssfff");   // Formato de fecha y hora
            return nombreCarpeta;
        }

        /// <summary>
        /// Agrega permisos de escritura a una carpeta.
        /// </summary>
        /// <param name="folderPath">La ruta de la carpeta a la que se le asignarán permisos de escritura.</param>
        static void AssignWritePrivilegesToDirectory(String folderPath)
        {
            DirectoryInfo newFileInfo = new DirectoryInfo(folderPath);               // Obtener información de la carpeta.
            DirectorySecurity newFileSecurity = newFileInfo.GetAccessControl();      // Obtener el control de acceso actual de la carpeta.
            FileSystemAccessRule writeRule = new FileSystemAccessRule(               // Crear una regla de acceso para permitir la escritura a todos los usuarios.
                new SecurityIdentifier(WellKnownSidType.BuiltinUsersSid, null),
                FileSystemRights.Write,
                InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit,
                PropagationFlags.None,
                AccessControlType.Allow
            );
            newFileSecurity.AddAccessRule(writeRule);                                // Agregar la regla de acceso al control de acceso.
            newFileInfo.SetAccessControl(newFileSecurity);                           // Establecer el nuevo control de acceso en la carpeta.
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

        /// <summary>
        /// Guarda y cierra un documento PDF.
        /// </summary>
        /// <param name="document">El objeto PdfDocument que se va a guardar y cerrar.</param>
        /// <param name="outputPath">La ruta de salida donde se guardará el documento PDF.</param>
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