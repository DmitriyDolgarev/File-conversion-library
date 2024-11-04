using System.Diagnostics;
using System.IO.Compression;
using PdfSharp.Pdf;
using PdfSharp.Drawing;
using PdfSharp.Pdf.IO;
using SkiaSharp;
using PDFtoImage;
//using Conversion = PDFtoImage.Conversion;

namespace FileConverterLib
{
    public class FileConverter
    {
        // Путь до soffice
        public static string sofficePath = @"C:\Program Files\LibreOffice\program";
        public static bool SofficeExists { get => File.Exists(Path.Combine(sofficePath, "soffice.exe")); }

        #region Merge PDFs
        public static void MergePDFs(string[] pdfFiles, string pdfOutput)
        {
            pdfOutput = Path.ChangeExtension(pdfOutput, "pdf");
            for (int i = 0; i < pdfFiles.Length; i++)
                pdfFiles[i] = Path.ChangeExtension(pdfFiles[i], "pdf");

            using (var outputDocument = new PdfDocument())
            {
                foreach (var pdf in pdfFiles)
                {
                    using (var inputDocument = PdfReader.Open(pdf, PdfDocumentOpenMode.Import))
                    {
                        for (int i = 0; i < inputDocument.PageCount; i++)
                        {
                            PdfPage page = inputDocument.Pages[i];
                            outputDocument.AddPage(page);
                        }
                    }
                }
                outputDocument.Save(pdfOutput);
            }
        }
        #endregion

        #region Split PDF
        public static void SplitPDF(string pdfInput, int pageSplitFrom, string pdf1Output, string pdf2Output)
        {
            pdfInput = Path.ChangeExtension(pdfInput, "pdf");
            pdf1Output = Path.ChangeExtension(pdf1Output, "pdf");
            pdf2Output = Path.ChangeExtension(pdf2Output, "pdf");

            using(var inputDocument = PdfReader.Open(pdfInput, PdfDocumentOpenMode.Import))
            {
                var outputDocument1 = new PdfDocument();
                var outputDocument2 = new PdfDocument();

                for(int i = 0; i < inputDocument.PageCount; i++)
                {
                    var page = inputDocument.Pages[i];

                    if (i < pageSplitFrom - 1)
                        outputDocument1.AddPage(page);
                    else
                        outputDocument2.AddPage(page);            
                }

                outputDocument1.Save(pdf1Output);
                outputDocument2.Save(pdf2Output);

                outputDocument1.Dispose();
                outputDocument2.Dispose();
            }
        }

        public static void SplitPDF(string pdfInput, int pageSplitFrom)
        {
            var pdf1 = GetFileNameInSameFolder(pdfInput) + "_splitted1.pdf";
            var pdf2 = GetFileNameInSameFolder(pdfInput) + "_splitted2.pdf";

            SplitPDF(pdfInput, pageSplitFrom, pdf1, pdf2);
        }
        #endregion

        #region JPG to PNG
        public static void JpgFileToPngFile(string jpgFileName, string pngFileName)
        {
            jpgFileName = Path.ChangeExtension(jpgFileName, "jpg");
            pngFileName = Path.ChangeExtension(pngFileName, "png");

            using(var img = SKImage.FromEncodedData(jpgFileName))
            {
                using(var data = img.Encode(SKEncodedImageFormat.Png, 100))
                {
                    using (var stream = File.OpenWrite(pngFileName))
                    {   
                        data.SaveTo(stream);
                    }
                }
            }
        }

        public static void JpgFileToPngFile(string jpgFileName)
        {
            JpgFileToPngFile(jpgFileName, GetFileNameInSameFolder(jpgFileName));
        }
        #endregion

        #region PNG TO JPG
        public static void PngFileToJpgFile(string pngFileName, string jpgFileName)
        {
            pngFileName = Path.ChangeExtension(pngFileName, "png");   
            jpgFileName = Path.ChangeExtension(jpgFileName, "jpg");

            using (var img = SKImage.FromEncodedData(pngFileName))
            {
                using (var data = img.Encode(SKEncodedImageFormat.Jpeg, 100))
                {
                    using (var stream = File.OpenWrite(jpgFileName))
                    {
                        data.SaveTo(stream);
                    }
                }
            }
        }

        public static void PngFileToJpgFile(string pngFileName)
        {
            PngFileToJpgFile(pngFileName, GetFileNameInSameFolder(pngFileName));
        }
        #endregion

        #region JPG to PDF
        public static void JpgFilesToPdfFile(string[] jpgFiles, string pdfFileName)
        {
            pdfFileName = Path.ChangeExtension(pdfFileName, "pdf");
            for (int i = 0; i < jpgFiles.Length; i++)
                jpgFiles[i] = Path.ChangeExtension(jpgFiles[i], "jpg");

            using (PdfDocument document = new PdfDocument())
            {
                foreach (var jpgFile in jpgFiles)
                {
                    PdfPage page = document.AddPage();

                    XGraphics gfx = XGraphics.FromPdfPage(page);
                    XImage image = XImage.FromFile(jpgFile);

                    page.Height = XUnit.FromPoint(image.PointHeight);
                    page.Width = XUnit.FromPoint(image.PointWidth);

                    gfx.DrawImage(image, 0, 0);
                }

                document.Save(pdfFileName);
            }
        }
        #endregion

        #region PDF to JPG
        public static void PdfFileToJpgFiles(string pdfFileName, string jpgFolderName, bool zip = false)
        {
            if (zip)
                PdfFileToJpgFilesZip(pdfFileName, jpgFolderName);
            else
                PdfFileToJpgFilesFolder(pdfFileName, jpgFolderName);
        }
        public static void PdfFileToJpgFiles(string pdfFileName, bool zip = false)
        {
            if (zip)
                PdfFileToJpgFilesZip(pdfFileName);
            else
                PdfFileToJpgFilesFolder(pdfFileName);
        }

        private static void PdfFileToJpgFilesFolder(string pdfFileName, string jpgFolderName)
        {
            pdfFileName = Path.ChangeExtension(pdfFileName, "pdf");
            if (!Directory.Exists(jpgFolderName))
                Directory.CreateDirectory(jpgFolderName);

            var pdfByteArray = Convert.ToBase64String(File.ReadAllBytes(pdfFileName));
            var pageCount = Conversion.GetPageCount(pdfByteArray);

            for (int i = 0; i < pageCount; i++)
            {
                var fileName = Path.Combine(jpgFolderName, $"page_{i + 1}.jpg");

                var pageSize = Conversion.GetPageSize(pdfByteArray, i);
                var options = new RenderOptions(
                            Dpi: 300,
                            Width: PointsToPixels(pageSize.Width),
                            Height: PointsToPixels(pageSize.Height)
                );

                using (var skImage = Conversion.ToImage(pdfByteArray, i, options: options))
                {
                    using (var data = skImage.Encode(SKEncodedImageFormat.Jpeg, 100))
                    {
                        using (var stream = File.OpenWrite(fileName))
                        {
                            data.SaveTo(stream);
                        }
                    }
                }
            }
        }

        private static void PdfFileToJpgFilesFolder(string pdfFileName)
        {
            PdfFileToJpgFiles(pdfFileName, GetFileNameInSameFolder(pdfFileName));
        }

        private static void PdfFileToJpgFilesZip(string pdfFileName, string jpgFolderName)
        {
            pdfFileName = Path.ChangeExtension(pdfFileName, "pdf");
            jpgFolderName = Path.ChangeExtension(jpgFolderName, "zip");


            using(var fs = new FileStream(jpgFolderName, FileMode.OpenOrCreate))
            {
                using (var archive = new ZipArchive(fs, ZipArchiveMode.Create, true))
                {
                    var pdfByteArray = Convert.ToBase64String(File.ReadAllBytes(pdfFileName));
                    var pageCount = Conversion.GetPageCount(pdfByteArray);

                    for (int i = 0; i < pageCount; i++)
                    {
                        var fileName = $"page_{i + 1}.jpg";

                        var pageSize = Conversion.GetPageSize(pdfByteArray, i);
                        var options = new RenderOptions(
                                    Dpi: 300,
                                    Width: PointsToPixels(pageSize.Width),
                                    Height: PointsToPixels(pageSize.Height)
                        );

                        using (var skImage = Conversion.ToImage(pdfByteArray, i, options: options))
                        {
                            using (var data = skImage.Encode(SKEncodedImageFormat.Jpeg, 100))
                            {
                                var archiveEntry = archive.CreateEntry(fileName);
                                using (var zipStream = archiveEntry.Open())
                                {
                                    using (var ms = new MemoryStream())
                                    {
                                        data.SaveTo(ms);
                                        zipStream.Write(ms.ToArray());
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        private static void PdfFileToJpgFilesZip(string pdfFileName)
        {
            PdfFileToJpgFilesZip(pdfFileName, GetFileNameInSameFolder(pdfFileName));
        }
        #endregion

        #region WORD to PDF
        public static void DocxFileToPdfFile(string wordFileName, string pdfFileFolder)
        {
            if(Path.GetExtension(wordFileName).ToLower() != ".docx" && Path.GetExtension(wordFileName).ToLower() != ".doc")
                wordFileName = Path.ChangeExtension(wordFileName, "docx");

            using (Process process = new Process())
            {
                ProcessStartInfo info = new ProcessStartInfo();
                info.FileName = "soffice";
                info.WorkingDirectory = sofficePath;
                info.Arguments = $"--headless --convert-to \"pdf:writer_pdf_Export\" \"{wordFileName}\" --outdir \"{pdfFileFolder}\"";
                info.UseShellExecute = true;
                process.StartInfo = info;

                process.Start();
                process.WaitForExit();
                process.Close();
            }
        }
        #endregion

        #region PDF to WORD
        public static void PdfFileToDocxFile(string pdfFileName, string wordFileFolder)
        {
            pdfFileName = Path.ChangeExtension(pdfFileName, "pdf");

            using (var pdfDocument = PdfReader.Open(pdfFileName))
            {
                
            }

            using (Process process = new Process())
            {
                ProcessStartInfo info = new ProcessStartInfo();
                info.FileName = "soffice";
                info.WorkingDirectory = sofficePath;
                info.Arguments = $"--headless --infilter=\"writer_pdf_import\" --convert-to docx \"{pdfFileName}\" --outdir \"{wordFileFolder}\"";
                info.UseShellExecute = true;
                process.StartInfo = info;

                process.Start();
                process.WaitForExit();
                process.Close();
            }
        }
        #endregion

        #region Pptx to PDF
        public static void PptxFileToPdfFile(string pptxFileName, string pdfFileFolder)
        {
            if (Path.GetExtension(pptxFileName).ToLower() != ".pptx" && Path.GetExtension(pptxFileName).ToLower() != ".ppt")
                pptxFileName = Path.ChangeExtension(pptxFileName, "pptx");
            using (Process process = new Process())
            {
                ProcessStartInfo info = new ProcessStartInfo();
                info.FileName = "soffice";
                info.WorkingDirectory = sofficePath;
                info.Arguments = $"--headless --convert-to \"pdf\" \"{pptxFileName}\" --outdir \"{pdfFileFolder}\"";
                info.UseShellExecute = true;
                process.StartInfo = info;

                process.Start();
                process.WaitForExit();
                process.Close();
            }
        }
        #endregion

        private static string GetFileNameInSameFolder(string fileName)
        {
            var dirPath = Path.GetDirectoryName(fileName);
            var _fileName = Path.GetFileNameWithoutExtension(fileName);

            return Path.Combine(dirPath, _fileName);
        }

        private static string GetFileNameInSameFolder(string fileName, string extension)
        {
            return GetFileNameInSameFolder(fileName) + "." + extension;
        }

        private static int PointsToPixels(double points)
        {
            return (int)Math.Round(points * 1.333f);
        }
    }
}
