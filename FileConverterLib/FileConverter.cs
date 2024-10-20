using System.Drawing;
using System.Drawing.Imaging;
using PdfSharp.Pdf;
using PdfSharp.Drawing;
using PdfSharp.Pdf.IO;
using System.Diagnostics;
using System;

namespace FileConverterLib
{
    public class FileConverter
    {
        // Путь до soffice
        public static string sofficePath = @"C:\Program Files\LibreOffice\program";
        public static bool SofficeExists { get => File.Exists(Path.Combine(sofficePath, "soffice.exe")); }

        #region Merge PDFs
        public static void MergePDFs(string pdfOutput, string[] pdfFiles)
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

            using (var img = Image.FromFile(jpgFileName))
            {
                img.Save(pngFileName, ImageFormat.Png);
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

            using (var img = Image.FromFile(pngFileName))
            {
                img.Save(jpgFileName, ImageFormat.Jpeg);
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
        public static void PdfFileToJpgFiles(string pdfFileName, string jpgFolderName)
        {
            pdfFileName = Path.ChangeExtension(pdfFileName, "pdf");
            if(!Directory.Exists(jpgFolderName))
                Directory.CreateDirectory(jpgFolderName);

            using(var pdfDocument = PdfiumViewer.PdfDocument.Load(pdfFileName))
            {
                for (int i = 0; i < pdfDocument.PageCount; i++)
                {
                    var bitmapImage = pdfDocument.Render(i, 300, 300, true);
                    var fileName = Path.Combine(jpgFolderName, $"page_{i + 1}.jpg");
                    bitmapImage.Save(fileName, ImageFormat.Jpeg);
                }
            }
        }

        public static void PdfFileToJpgFiles(string pdfFileName)
        {
            PdfFileToJpgFiles(pdfFileName, GetFileNameInSameFolder(pdfFileName));
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
    }
}
