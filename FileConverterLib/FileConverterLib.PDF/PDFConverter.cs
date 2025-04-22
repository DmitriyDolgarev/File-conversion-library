using System.Diagnostics;
using System.IO.Compression;
using PdfSharp.Pdf;
using PdfSharp.Drawing;
using PdfSharp.Pdf.IO;
using SkiaSharp;
using PDFtoImage;
using FileConverterLib.Utils;
using System.Reflection.Metadata;

namespace FileConverterLib.PDF
{
    public class PdfConverter
    {
        private static int PointsToPixels(double points)
        {
            return (int)Math.Round(points * 1.333f);
        }

        #region Merge PDFs
        public static void MergePdfFiles(string[] pdfFiles, string pdfOutput)
        {
            pdfOutput = FileConverterUtils.GetCorrectedPath(pdfOutput, "pdf");
            for (int i = 0; i < pdfFiles.Length; i++)
                pdfFiles[i] = FileConverterUtils.GetCorrectedPath(pdfFiles[i], "pdf");

            var pdfDocuments = new List<PdfDocument>();
            foreach(var pdf in pdfFiles)
            {
                var document = PdfReader.Open(pdf, PdfDocumentOpenMode.Import);
                pdfDocuments.Add(document);
            }

            using var outputDocument = MergePdfDocuments(pdfDocuments);
            outputDocument.Save(pdfOutput);

            pdfDocuments.ForEach(doc => doc.Dispose());
        }
        public static byte[] MergePdfBytes(IEnumerable<byte[]> pdfFiles)
        {
            var pdfDocuments = new List<PdfDocument>();
            foreach(var pdf in pdfFiles)
            {
                using var pdfStream = new MemoryStream(pdf);
                var document = PdfReader.Open(pdfStream, PdfDocumentOpenMode.Import);
                pdfDocuments.Add(document);
            }

            using var outputDocument = MergePdfDocuments(pdfDocuments);

            using (var stream = new MemoryStream())
            {
                outputDocument.Save(stream);

                pdfDocuments.ForEach(doc => doc.Dispose());
                outputDocument.Dispose();

                return stream.ToArray();
            }
        }
        private static PdfDocument MergePdfDocuments(IEnumerable<PdfDocument> pdfDocuments)
        {
            var outputDocument = new PdfDocument();

            foreach (var pdf in pdfDocuments)
            {
                for (int i = 0; i < pdf.PageCount; i++)
                {
                    PdfPage page = pdf.Pages[i];
                    outputDocument.AddPage(page);
                }
            }

            return outputDocument;
        }
        #endregion

        #region Split PDF
        public static void SplitPdfFile(string pdfInput, int pageSplitFrom, string pdf1Output, string pdf2Output)
        {
            pdfInput = FileConverterUtils.GetCorrectedPath(pdfInput, "pdf");
            pdf1Output = FileConverterUtils.GetCorrectedPath(pdf1Output, "pdf");
            pdf2Output = FileConverterUtils.GetCorrectedPath(pdf2Output, "pdf");

            using (var inputDocument = PdfReader.Open(pdfInput, PdfDocumentOpenMode.Import))
            {
                var outputDocument1 = new PdfDocument();
                var outputDocument2 = new PdfDocument();

                for (int i = 0; i < inputDocument.PageCount; i++)
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
        public static void SplitPdfFile(string pdfInput, int pageSplitFrom)
        {
            var pdf1 = FileConverterUtils.GetFileNameInSameFolder(pdfInput) + "_splitted1.pdf";
            var pdf2 = FileConverterUtils.GetFileNameInSameFolder(pdfInput) + "_splitted2.pdf";

            SplitPdfFile(pdfInput, pageSplitFrom, pdf1, pdf2);
        }
        #endregion

        #region JPG to PDF
        public static void JpgFilesToPdfFile(string[] jpgFiles, string pdfFileName)
        {
            pdfFileName = FileConverterUtils.GetCorrectedPath(pdfFileName, "pdf");
            for (int i = 0; i < jpgFiles.Length; i++)
                jpgFiles[i] = FileConverterUtils.GetCorrectedPath(jpgFiles[i], "jpg");

            var jpgBytes = new List<byte[]>();

            foreach(var jpgFile in jpgFiles)
            {
                jpgBytes.Add(File.ReadAllBytes(jpgFile));
            }

            using var pdfDocument = JpgToPdf(jpgBytes);
            pdfDocument.Save(pdfFileName);
        }
        public static byte[] JpgBytesToPdfBytes(IEnumerable<byte[]> jpgBytes)
        {
            using var pdfDocument = JpgToPdf(jpgBytes);

            using (var stream = new MemoryStream())
            {
                pdfDocument.Save(stream);
                return stream.ToArray();
            }
        }
        private static PdfDocument JpgToPdf(IEnumerable<byte[]> jpgBytes)
        {
            var pdfDocument = new PdfDocument();

            foreach (var jpg in jpgBytes)
            {
                using (var stream = new MemoryStream())
                {
                    stream.Write(jpg);
                    PdfPage page = pdfDocument.AddPage();

                    XGraphics gfx = XGraphics.FromPdfPage(page);
                    XImage image = XImage.FromStream(stream);

                    page.Height = XUnit.FromPoint(image.PointHeight);
                    page.Width = XUnit.FromPoint(image.PointWidth);

                    gfx.DrawImage(image, 0, 0);
                }
            }

            return pdfDocument;
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
            pdfFileName = FileConverterUtils.GetCorrectedPath(pdfFileName, "pdf");
            jpgFolderName = Path.GetFullPath(jpgFolderName);
            if (!Directory.Exists(jpgFolderName))
                Directory.CreateDirectory(jpgFolderName);

            var pdfByteArray = File.ReadAllBytes(pdfFileName);

            var jpgBytesList = PdfBytesToJpgBytes(pdfByteArray);

            for(int i = 0; i < jpgBytesList.Count; i++)
            {
                var jpgBytes = jpgBytesList[i];
                var fileName = Path.Combine(jpgFolderName, $"page_{i + 1}.jpg");

                using (var stream = File.OpenWrite(fileName))
                {
                    stream.Write(jpgBytes);
                }
            }
        }
        private static void PdfFileToJpgFilesFolder(string pdfFileName)
        {
            PdfFileToJpgFiles(pdfFileName, FileConverterUtils.GetFileNameInSameFolder(pdfFileName));
        }
        private static void PdfFileToJpgFilesZip(string pdfFileName, string jpgFolderName)
        {
            pdfFileName = FileConverterUtils.GetCorrectedPath(pdfFileName, "pdf");
            jpgFolderName = FileConverterUtils.GetCorrectedPath(jpgFolderName, "zip");

            var pdfByteArray = File.ReadAllBytes(pdfFileName);

            var zipBytes = PdfBytesToJpgBytesZip(pdfByteArray);

            using(var fs = new FileStream(jpgFolderName, FileMode.OpenOrCreate))
            {
                fs.Write(zipBytes);
            }
        }
        private static void PdfFileToJpgFilesZip(string pdfFileName)
        {
            PdfFileToJpgFilesZip(pdfFileName, FileConverterUtils.GetFileNameInSameFolder(pdfFileName));
        }
        
        public static List<byte[]> PdfBytesToJpgBytes(byte[] pdfBytes)
        {
            var jpgBytes = new List<byte[]>();

            var pageCount = Conversion.GetPageCount(pdfBytes);

            for (int i = 0; i < pageCount; i++)
            {
                var pageSize = Conversion.GetPageSize(pdfBytes, i);
                var options = new RenderOptions(
                            Dpi: 300,
                            Width: PointsToPixels(pageSize.Width),
                            Height: PointsToPixels(pageSize.Height)
                );

                using (var skImage = Conversion.ToImage(pdfBytes, i, options: options))
                {
                    using (var data = skImage.Encode(SKEncodedImageFormat.Jpeg, 100))
                    {
                        using (var stream = new MemoryStream())
                        {
                            data.SaveTo(stream);
                            jpgBytes.Add(stream.ToArray());
                        }
                    }
                }
            }

            return jpgBytes;
        }
        public static byte[] PdfBytesToJpgBytesZip(byte[] pdfBytes)
        {
            var jpgBytesList = PdfBytesToJpgBytes(pdfBytes);

            using(var stream = new MemoryStream())
            {
                using (var archive = new ZipArchive(stream, ZipArchiveMode.Create))
                {
                    for(int i = 0; i < jpgBytesList.Count; i++)
                    {
                        var jpgBytes = jpgBytesList[i];
                        var archiveEntry = archive.CreateEntry($"page_{i+1}.jpg");
                        
                        using(var zipStream = archiveEntry.Open())
                        {
                            zipStream.Write(jpgBytes);
                        }
                    }
                }

                return stream.ToArray();
            }
        }
        #endregion
    }
}
