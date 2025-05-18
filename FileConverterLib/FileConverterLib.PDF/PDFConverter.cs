using System.IO.Compression;
using PdfSharp.Pdf;
using PdfSharp.Drawing;
using PdfSharp.Pdf.IO;
using SkiaSharp;
using PDFtoImage;
using FileConverterLib.Utils;
using System.Text.RegularExpressions;

namespace FileConverterLib.PDF
{
    public partial class PdfConverter
    {
        private static int PointsToPixels(double points)
        {
            return (int)Math.Round(points * 1.333f);
        }
        private static IEnumerable<int> GetPagesList(string pagesStr)
        {

            var pages = new List<int>();
            var correctNumber = @"^\d+$";

            var splitPagesStr = pagesStr.Split([' ', ',', ';']);

            foreach (var page in splitPagesStr)
            {
                var p = page.Split('-');
                Console.WriteLine(p[0]);
                if (p.Length > 1) // Range
                {
                    if (!(Regex.IsMatch(p[0], correctNumber) && Regex.IsMatch(p[1], correctNumber)))
                        throw new ArgumentException("Invalid number format");

                    var start = int.Parse(p[0]);
                    var end = int.Parse(p[1]);

                    if (start < end)
                    {
                        for (var i = start; i <= end; i++)
                            pages.Add(i);
                    }
                    else
                    {
                        for (var i = start; i >= end; i--)
                            pages.Add(i);
                    }
                }
                else // Single page
                {
                    if (Regex.IsMatch(p[0], correctNumber))
                        pages.Add(int.Parse(p[0]));
                }
            }

            return pages;
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

        public static void SplitPdfFile(string pdfInput, string splitString, string pdfOutput)
        {
            pdfInput = FileConverterUtils.GetCorrectedPath(pdfInput, "pdf");
            pdfOutput = FileConverterUtils.GetCorrectedPath(pdfOutput, "pdf");

            using (var inputDocument = PdfReader.Open(pdfInput, PdfDocumentOpenMode.Import))
            {
                using(var outputDocument = SplitPdf(inputDocument, splitString))
                {
                    outputDocument.Save(pdfOutput);
                }
            }
        }
        public static void SplitPdfFile(string pdfInput, string splitString)
        {
            var pdf = FileConverterUtils.GetFileNameInSameFolder(pdfInput) + "_splitted.pdf";
            SplitPdfFile(pdfInput, splitString, pdf);
        }
        public static byte[] SplitPdfBytes(byte[] pdfBytes, string splitString)
        {
            using var pdfStream = new MemoryStream(pdfBytes);
            using var inputDocument = PdfReader.Open(pdfStream, PdfDocumentOpenMode.Import);
            using var outputDocument = SplitPdf(inputDocument, splitString);

            using var stream = new MemoryStream();
            outputDocument.Save(stream);
            return stream.ToArray();
        }
        private static PdfDocument SplitPdf(PdfDocument inputDocument, string splitString)
        {
            var outputDocument = new PdfDocument();
            var pages = GetPagesList(splitString);

            foreach (var pageNum in pages)
            {
                var page = inputDocument.Pages[pageNum - 1];
                outputDocument.AddPage(page);
            }

            return outputDocument;
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

            int page = 1;
            foreach(var jpgBytes in jpgBytesList)
            {
                var fileName = Path.Combine(jpgFolderName, $"page_{page}.jpg");

                using (var stream = File.OpenWrite(fileName))
                {
                    stream.Write(jpgBytes);
                }

                page++;
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
        
        public static IEnumerable<byte[]> PdfBytesToJpgBytes(byte[] pdfBytes)
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
                    int page = 1;
                    foreach(var jpgBytes in jpgBytesList)
                    {
                        var archiveEntry = archive.CreateEntry($"page_{page}.jpg");
                        
                        using(var zipStream = archiveEntry.Open())
                        {
                            zipStream.Write(jpgBytes);
                        }
                        page++;
                    }
                }

                return stream.ToArray();
            }
        }
        #endregion
    }
}
