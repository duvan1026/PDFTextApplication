using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Security.AccessControl;
using System.Security.Principal;
using PdfSharp.Drawing;
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
        const string nameDirectoryDestination = "Data.Process";
        const string inputFormat = "*.tif";
        const string outputFormat = ".pdf";

        #endregion

        static void Main()
        {
            // Iniciar el cronómetro para medir el tiempo de ejecución
            Stopwatch stopwatch = Stopwatch.StartNew();

            // Crear el directorio de destino para el archivo de salida
            string outputFileDestination = Path.Combine(outputFile, nameDirectoryDestination);      
            CreateDirectoryWithWriteAccess(outputFileDestination);

            // Generar un nombre de directorio basado en la fecha y hora actual
            string outputDirectoryDate = Path.Combine(outputFileDestination, GetNameFileDate());
            CreateDirectoryWithWriteAccess(outputDirectoryDate);

            ProcessInputDirectories(inputFile, outputDirectoryDate);

            // Detener el cronómetro y calcular el tiempo transcurrido
            stopwatch.Stop(); 
            TimeSpan elapsedTime = stopwatch.Elapsed;
            Console.WriteLine($"Proceso completado en {elapsedTime.TotalMinutes} minutos {elapsedTime.Seconds} segundos");
            Console.WriteLine("Proceso completado. El PDF de texto se ha guardado.");
            Console.ReadLine();
        }

        /// <summary>
        /// Procesa los directorios de entrada y mueve archivos a directorios de destino basados en la fecha.
        /// </summary>
        /// <param name="inputFile">El directorio de entrada que contiene subdirectorios a procesar.</param>
        /// <param name="outputDirectoryDate">El directorio de destino basado en la fecha donde se moverán los archivos.</param>
        static void ProcessInputDirectories(string inputFile, string outputDirectoryDate)
        {
            // Recorrer los directorios de entrada
            foreach (string directoryInput in Directory.GetDirectories(inputFile))
            {
                string nameDirectoryInput = Path.GetFileName(directoryInput);

                // Verificar si el directorio actual no es el directorio de destino
                if (nameDirectoryInput != nameDirectoryDestination)
                {
                    // Crear la ruta de destino para el directorio actual
                    string destinationRoute = Path.Combine(outputDirectoryDate, nameDirectoryInput);
                    CreateDirectoryWithWriteAccess(destinationRoute);

                    ProcessTiffFiles(directoryInput, destinationRoute);
                }
            }
        }

        /// <summary>
        /// Procesa archivos TIFF en un directorio de entrada y los convierte en un documento PDF de destino.
        /// </summary>
        /// <param name="directoryInput">El directorio de entrada que contiene archivos TIFF a procesar.</param>
        /// <param name="destinationRoute">El directorio de destino donde se almacenará el documento PDF resultante.</param>
        static void ProcessTiffFiles(string directoryInput, string destinationRoute)
        {
            // Obtener archivos TIFF en el directorio actual
            string[] tiffFiles = GetTiffFilesInDirectory(directoryInput);
            if (tiffFiles.Length == 0) return;

            string outputnamePDFTotal = GetOutputPdfName(destinationRoute);

            using (PdfDocument outputDocument = CreatePdfDocument())
            {
                // Procesar cada archivo TIFF y convertirlo a PDF
                foreach (string tiffFile in tiffFiles)
                {
                    string outputPath = GetPdfOutputPath(destinationRoute, tiffFile);
                    ConvertTiffToPdf(tiffFile, outputPath);
                    AddPagesFromTiffToPdf(outputPath, outputDocument);
                }
                SavePdfDocument(outputDocument, outputnamePDFTotal);
            }
        }

        /// <summary>
        /// Obtiene el nombre del archivo PDF de salida basado en el directorio de destino.
        /// </summary>
        /// <param name="destinationRoute">El directorio de destino para el archivo PDF.</param>
        /// <returns>El nombre del archivo PDF de salida.</returns>
        static string GetOutputPdfName(string destinationRoute)
        {
            return Path.Combine(destinationRoute, Path.GetFileName(destinationRoute) + outputFormat);
        }

        /// <summary>
        /// Obtiene archivos TIFF en un directorio de entrada.
        /// </summary>
        /// <param name="directoryInput">El directorio de entrada que contiene archivos TIFF.</param>
        /// <returns>Un arreglo de rutas de archivo TIFF encontrados en el directorio de entrada.</returns>
        static string[] GetTiffFilesInDirectory(string directoryInput)
        {
            return Directory.GetFiles(directoryInput, inputFormat);
        }

        /// <summary>
        /// Crea un nuevo documento PDF.
        /// </summary>
        /// <returns>Un objeto PdfDocument que se utilizará como documento PDF de destino.</returns>
        static PdfDocument CreatePdfDocument()
        {
            return new PdfDocument();
        }

        /// <summary>
        /// Obtiene la ruta de salida del archivo PDF basada en la ruta de destino y el archivo TIFF de entrada.
        /// </summary>
        /// <param name="destinationRoute">La ruta de destino para el archivo PDF.</param>
        /// <param name="tiffFile">La ruta del archivo TIFF de entrada.</param>
        /// <returns>La ruta del archivo PDF de salida.</returns>
        static string GetPdfOutputPath(string destinationRoute, string tiffFile)
        {
            return Path.Combine(destinationRoute, Path.GetFileNameWithoutExtension(tiffFile) + outputFormat);
        }

        /// <summary>
        /// Agrega las páginas de un archivo TIFF a un documento PDF de destino.
        /// </summary>
        /// <param name="tiffFilePath">La ruta del archivo TIFF del cual se extraerán las páginas.</param>
        /// <param name="targetDocument">El documento PDF de destino al cual se agregarán las páginas del archivo TIFF.</param>
        static void AddPagesFromTiffToPdf(string tiffFilePath, PdfDocument targetDocument)
        {
            using (PdfDocument inputDocument = PdfReader.Open(tiffFilePath, PdfDocumentOpenMode.Import))
            {
                int pageCounter = inputDocument.PageCount;
                for (int pageIndex = 0; pageIndex < pageCounter; pageIndex++)
                {
                    PdfPage page = inputDocument.Pages[pageIndex];
                    targetDocument.AddPage(page);
                }
            }
        }



        // TODO: Arregloarlo para implementarlo en la libreriaOCR
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

                        var imageSize = new XSize(ConvertToPoints(image.Width, dpi), 
                                                  ConvertToPoints(image.Height, dpi));

                        using (var document = CreateCustomPdfDoc(imageSize))
                        {
                            using (var gfx = CreateGraphics(document))
                            {
                                scaleWidth = CalculateScale(imageSize.Width, image.Width);
                                scaleHeight = CalculateScale(imageSize.Height, image.Height);

                                AddImageToPdf(gfx, tiffImagePath, imageSize);
                                ProcessText(pageProcessor, gfx, scaleWidth, scaleHeight);
                                SavePdfDocument(document, pdfOutputPath);
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

        /// <summary>
        /// Calcula la escala entre dos valores numéricos.
        /// </summary>
        /// <param name="value1">El primer valor numérico.</param>
        /// <param name="value2">El segundo valor numérico (divisor).</param>
        /// <returns>El resultado de dividir el primer valor por el segundo valor, que representa la escala.</returns>
        static double CalculateScale(double value1, double value2)
        {
            return value1 / value2;
        }

        /// <summary>
        /// Crea un documento PDF personalizado con las dimensiones especificadas.
        /// </summary>
        /// <param name="imageSize">El tamaño del documento PDF en formato XSize (ancho y alto).</param>
        /// <returns>Un objeto PdfDocument que representa el nuevo documento PDF con las dimensiones especificadas.</returns>
        static PdfDocument CreateCustomPdfDoc(XSize imageSize)
        {
            var document = CreatePdfDocument();
            var page = document.AddPage();
            page.Width = imageSize.Width;
            page.Height = imageSize.Height;
            return document;
        }

        /// <summary>
        /// Guarda un documento PDF.
        /// </summary>
        /// <param name="document">El objeto PdfDocument que se va a guardar y cerrar.</param>
        /// <param name="outputPath">La ruta de salida donde se guardará el documento PDF.</param>
        static void SavePdfDocument(PdfDocument document, string outputPath)
        {
            document.Save(outputPath);
        }

        /// <summary>
        /// Crea un contexto gráfico (XGraphics) para dibujar en la primera página de un documento PDF.
        /// </summary>
        /// <param name="document">El documento PDF en el que se va a crear el contexto gráfico.</param>
        /// <returns>Un objeto XGraphics que proporciona un contexto gráfico para dibujar en la primera página del documento PDF.</returns>
        static XGraphics CreateGraphics(PdfDocument document)
        {
            var page = document.Pages[0];
            return XGraphics.FromPdfPage(page);
        }

        /// <summary>
        /// Convierte un valor de longitud desde una unidad de medida específica a puntos (72 puntos por pulgada) en un contexto de resolución dado.
        /// </summary>
        /// <param name="value">El valor de longitud que se desea convertir.</param>
        /// <param name="dpi">La resolución en puntos por pulgada (DPI) en la que se realiza la conversión.</param>
        /// <returns>El valor de longitud convertido a puntos en el contexto de resolución especificado.</returns>
        static double ConvertToPoints(double value, double dpi)
        {
            return value * 72.0 / dpi;
        }

        /// <summary>
        /// Procesa y dibuja palabras en un documento PDF.
        /// </summary>
        /// <param name="pageProcessor">El procesador de páginas Tesseract.</param>
        /// <param name="gfx">El contexto gráfico para el PDF.</param>
        /// <param name="scaleWidth">Factor de escala para el ancho.</param>
        /// <param name="scaleHeight">Factor de escala para la altura.</param>
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

                        double realSizeInPoints = CalculateFontSize(bounds.Height);

                        DrawTextOnPdf(gfx, word, new XPoint(x1, y1), realSizeInPoints);
                    }
                }
            }
        }

        /// <summary>
        /// Dibuja el texto en un documento PDF en la ubicación especificada con el tamaño de fuente indicado.
        /// </summary>
        /// <param name="gfx">El contexto gráfico para el PDF.</param>
        /// <param name="text">El texto a dibujar.</param>
        /// <param name="position">La posición (X, Y) en la que se dibujará el texto.</param>
        /// <param name="fontSize">Tamaño de fuente en puntos.</param>
        static void DrawTextOnPdf(XGraphics gfx, string text, XPoint position, double fontSize)
        {
            var font = new XFont("Arial", fontSize);                // Agrega la palabra con el tamaño de fuente calculado
            XBrush brush = XBrushes.Transparent;                    // Establecer el color del texto como transparente
            gfx.DrawString(text, font, brush, position);    // Agregar la palabra y sus coordenadas al PDF
            //gfx.DrawString(word, font, XBrushes.Black, new XPoint(x1, y1));  // Agregar la palabra y sus coordenadas al PDF
        }


        /// <summary>
        /// Agrega una imagen desde un archivo a un documento PDF en el contexto gráfico especificado.
        /// </summary>
        /// <param name="gfx">El contexto gráfico donde se agregará la imagen al PDF.</param>
        /// <param name="imagePath">La ruta del archivo de imagen a agregar al PDF.</param>
        /// <param name="imageSize">El tamaño de la imagen a agregar en el formato XSize (ancho y alto).</param>
        static void AddImageToPdf(XGraphics gfx, string imagePath, XSize imageSize)
        {
            var xImage = XImage.FromFile(imagePath);
            gfx.DrawImage(xImage, 0, 0, imageSize.Width, imageSize.Height);
        }

        /// <summary>
        /// Calcula el tamaño de fuente en puntos para que la altura de la fuente coincida con el valor deseado en píxeles.
        /// </summary>
        /// <param name="targetHeightInPixels">La altura de la fuente deseada en píxeles.</param>
        /// <returns>El tamaño de fuente en puntos que produce la altura de fuente deseada.</returns>
        static double CalculateFontSize( double targetHeightInPixels)
        {
            const double tolerance = 1;
            const string fontFamilyName = "Arial";
            double realSizeInPoints = 12;

            XFont fontTest = new XFont(fontFamilyName, realSizeInPoints);
            double fontHeightInPixels = fontTest.GetHeight();

            while (Math.Abs(fontHeightInPixels - targetHeightInPixels) > tolerance)
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