
using System;
//using System.Linq;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.AccessControl;
using System.Security.Principal;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using TesseractOCR.Library.src;

namespace PDFTextApplication
{
    class Program
    {
        #region Constantes

        const string tessdataPath = @"C:\Program Files (x86)\Tesseract-OCR\tessdata";
        const string language = "spa";
        const double dpi = 96.0;                                  // Resolución estándar de pantalla
        
        //const string inputFile = @"C:\DuvanCastro\AplicacionAAA\Data"; // Reemplaza con la ruta de tu carpeta
        //const string outputFile = @"C:\DuvanCastro\AplicacionAAA\Data";

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
            //// Recorrer los directorios de entrada
            //foreach (string directoryInput in Directory.GetDirectories(inputFile))
            //{

            DirectoryInfo inputDirectory = new DirectoryInfo(inputFile);

            // Filtrar y recorrer solo los directorios que cumplen con la condición
            foreach (var directoryInput in inputDirectory.GetDirectories().Where(dir => !dir.Name.EndsWith(".#")))
            {
                string nameDirectoryInput = Path.GetFileName(directoryInput.FullName);

                // Verificar si el directorio actual no es el directorio de destino
                if (nameDirectoryInput != nameDirectoryDestination)
                {
                    // Crear la ruta de destino para el directorio actual
                    string destinationRoute = Path.Combine(outputDirectoryDate, nameDirectoryInput);
                    CreateDirectoryWithWriteAccess(destinationRoute);

                    ProcessTiffFiles(directoryInput.FullName, destinationRoute);
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
            //PDFService pDFService = new PDFService();

            // Obtener archivos TIFF en el directorio actual
            string[] tiffFiles = GetTiffFilesInDirectory(directoryInput);
            if (tiffFiles == null || tiffFiles.Length == 0) return;

            string outputnamePDFTotal = GetOutputPdfName(destinationRoute);

            using (PdfDocument outputDocument = CreatePdfDocument())
            {
                // Procesar cada archivo TIFF y convertirlo a PDF
                foreach (string tiffFile in tiffFiles)
                {
                    string outputPath = GetPdfOutputPath(destinationRoute, tiffFile);
                    PDFService.ConvertTiffToPdf(tiffFile, outputPath);
                    AddPagesFromTiffToPdf(outputPath, outputDocument);
                }
                SavePdfDocument(outputDocument, outputnamePDFTotal);
                // TODO: Cambiar nombre d ela carpeta  aña cual se extrajo la informacion
            }

            AgregarSufijo("C:\\Users\\duvan.castro\\Desktop\\TestPDFText\\Data\\testdata", ".#");
            AgregarSufijo(directoryInput, ".#");
        }

        static void AgregarSufijo(string rutaOriginal, string sufijo)
        {
            try
            {
                // Obtener el nombre de la carpeta actual
                string nombreCarpeta = Path.GetFileName(rutaOriginal);

                // Agregar el sufijo al nombre
                string nuevoNombre = String.Concat(nombreCarpeta, sufijo);

                // Obtener la ruta del directorio padre
                string directorioPadre = Path.GetDirectoryName(rutaOriginal);

                // Combinar la ruta del directorio padre con el nuevo nombre
                string nuevaRuta = Path.Combine(directorioPadre, nuevoNombre);

                AssignWritePrivilegesToDirectory(rutaOriginal);
                // Cambiar el nombre de la carpeta
                string newFolderPath = rutaOriginal + ".#";
                ChangeFolderName(rutaOriginal, newFolderPath);




                // Obtener el nombre del usuario de Windows actual
                string currentUserName = WindowsIdentity.GetCurrent().Name;
                Console.WriteLine("Usuario de Windows actual: " + currentUserName);

                //if (IsAdministrator())
                //{
                //    // Lógica de la aplicación
                //    // ...
                //    Console.WriteLine("Soy administrador");
                //    // Verificar los permisos de escritura en la carpeta
                //    if (HasWritePermissions(rutaOriginal))
                //    {
                //        Console.WriteLine("El usuario tiene permisos de escritura en la carpeta.");

                //        // Cambiar el nombre de la carpeta
                //        //string newFolderPath = rutaOriginal + ".#";
                //        //ChangeFolderName(rutaOriginal, newFolderPath);
                //    }
                //    else
                //    {
                //        Console.WriteLine("El usuario NO tiene permisos de escritura en la carpeta.");
                //    }
                //}
                //else
                //{
                //    // Si no se ejecuta como administrador, relanzar la aplicación con permisos elevados
                //    RunElevated();
                //}





                //using (var stream = new FileStream(rutaOriginal, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                //{
                //    // La carpeta no está en uso, puedes proceder a realizar operaciones
                //    Console.WriteLine("La carpeta no está en uso.");

                //    // Ahora puedes cambiar el nombre de la carpeta u realizar otras operaciones
                //    //CambiarNombreCarpeta(rutaCarpeta, "nuevo_nombre");
                //}


                //// Obtenemos una referencia a la carpeta
                //var directory = new DirectoryInfo(rutaOriginal);

                //// Establecemos el atributo ReadOnly a False
                //directory.Attributes &= ~FileAttributes.ReadOnly;

                ////if (VerificarAtributoSoloLectura(rutaOriginal))
                ////{
                ////    EliminarAtributoSoloLectura(rutaOriginal);
                ////}

                //// Mover la carpeta con el nuevo nombre
                //Directory.Move(rutaOriginal, nuevaRuta);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error al agregar sufijo: {ex.Message}");
            }
        }

        static bool IsAdministrator()
        {
            WindowsIdentity identity = WindowsIdentity.GetCurrent();
            WindowsPrincipal principal = new WindowsPrincipal(identity);
            return principal.IsInRole(WindowsBuiltInRole.Administrator);
        }

        static void RunElevated()
        {
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.UseShellExecute = true;
            startInfo.WorkingDirectory = Environment.CurrentDirectory;
            startInfo.FileName = Assembly.GetEntryAssembly().CodeBase;
            startInfo.Verb = "runas"; // Esto solicitará elevación de privilegios

            try
            {
                Process.Start(startInfo);
            }
            catch (System.ComponentModel.Win32Exception)
            {
                // El usuario canceló la solicitud de permisos elevados
            }
        }

        static void ChangeFolderName(string oldFolderPath, string newFolderPath)
        {
            try
            {
                // Cambiar el nombre de la carpeta
                System.IO.Directory.Move(oldFolderPath, newFolderPath);
            }
            catch (System.IO.IOException ex)
            {
                // Verificar si la excepción es debido a que la carpeta está en uso por otro proceso
                if (IsFolderInUse(ex))
                {
                    Console.WriteLine("La carpeta está siendo utilizada por otro proceso.");
                }
                else
                {
                    Console.WriteLine($"Error al cambiar el nombre de la carpeta: {ex.Message}");
                }
            }
            catch (UnauthorizedAccessException ex)
            {
                Console.WriteLine($"Error de acceso no autorizado al cambiar el nombre de la carpeta: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error al cambiar el nombre de la carpeta: {ex.Message}");
            }
        }

        static bool IsFolderInUse(System.IO.IOException ex)
        {
            int errorCode = System.Runtime.InteropServices.Marshal.GetHRForException(ex) & ((1 << 16) - 1);
            return errorCode == 32 || errorCode == 33; // 32: El proceso no puede obtener acceso al archivo porque está siendo utilizado por otro proceso, 33: El proceso no puede obtener acceso al archivo porque otro proceso tiene bloqueado una porción del archivo.
        }

        static bool HasWritePermissions(string folderPath)
        {
            try
            {
                // Obtener los permisos de la carpeta
                DirectorySecurity directorySecurity = System.IO.Directory.GetAccessControl(folderPath);

                // Obtener la identidad del usuario actual
                WindowsIdentity windowsIdentity = WindowsIdentity.GetCurrent();
                WindowsPrincipal windowsPrincipal = new WindowsPrincipal(windowsIdentity);

                // Verificar si el usuario actual tiene permisos de escritura
                AuthorizationRuleCollection rules = directorySecurity.GetAccessRules(true, true, typeof(SecurityIdentifier));
                foreach (FileSystemAccessRule rule in rules)
                {
                    if (windowsPrincipal.IsInRole(rule.IdentityReference as SecurityIdentifier) &&
                        (rule.FileSystemRights & FileSystemRights.WriteData) == FileSystemRights.WriteData)
                    {
                        return true;
                    }
                }

                return false;
            }
            catch (Exception ex)
            {
                // Manejar la excepción según tus necesidades
                Console.WriteLine("Error al verificar permisos: " + ex.Message);
                return false;
            }
        }

        static bool VerificarAtributoSoloLectura(string rutaCarpeta)
        {
            try
            {
                // Verificar si el atributo de solo lectura está presente
                return (File.GetAttributes(rutaCarpeta) & FileAttributes.ReadOnly) == FileAttributes.ReadOnly;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error al verificar el atributo de solo lectura: {ex.Message}");
                return false;
            }
        }

        static void EliminarAtributoSoloLectura(string rutaCarpeta)
        {
            try
            {
                // Crear un objeto DirectoryInfo para obtener información de la carpeta
                DirectoryInfo directorioInfo = new DirectoryInfo(rutaCarpeta);

                // Eliminar el atributo de solo lectura si está presente
                if ((directorioInfo.Attributes & FileAttributes.ReadOnly) == FileAttributes.ReadOnly)
                {
                    directorioInfo.Attributes &= ~FileAttributes.ReadOnly;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error al eliminar el atributo de solo lectura: {ex.Message}");
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
        /// Guarda un documento PDF.
        /// </summary>
        /// <param name="document">El objeto PdfDocument que se va a guardar y cerrar.</param>
        /// <param name="outputPath">La ruta de salida donde se guardará el documento PDF.</param>
        static void SavePdfDocument(PdfDocument document, string outputPath)
        {
            document.Save(outputPath);
        }   
        
    }
}