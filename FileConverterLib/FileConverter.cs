using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.Advanced;
using PdfSharp.Pdf.IO;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.ExceptionServices;

namespace FileConverterLib
{
    public class FileConverter
    {
        #region Merge PDFs
        public static void MergePDFs(string pdfOutput, params string[] pdfs)
        {
            using (var outputDocument = new PdfDocument())
            {
                foreach (var pdf in pdfs)
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
        #endregion

        #region JPG to PNG
        public static void JpgFileToPngFile(string jpgFileName, string pngFileName)
        {
            if (Path.GetExtension(jpgFileName) == "")
                jpgFileName += ".jpg";
            if (Path.GetExtension(pngFileName) == "")
                pngFileName += ".png";

            using (var img = Image.FromFile(jpgFileName))
            {
                img.Save(pngFileName, ImageFormat.Png);
            }
        }

        public static void JpgFileToPngFile(string jpgFileName)
        {
            string pngFileName = $@"{Path.GetDirectoryName(jpgFileName)}\{Path.GetFileNameWithoutExtension(jpgFileName)}.png";
            JpgFileToPngFile(jpgFileName, pngFileName);
        }
        #endregion

        #region PNG TO JPG
        public static void PngFileToJpgFile(string pngFileName, string jpgFileName)
        {
            if (Path.GetExtension(pngFileName) == "")
                pngFileName += ".png";
            if (Path.GetExtension(jpgFileName) == "")
                jpgFileName += ".jpg";

            using (var img = Image.FromFile(pngFileName))
            {
                img.Save(jpgFileName, ImageFormat.Jpeg);
            }
        }

        public static void PngFileToJpgFile(string pngFileName)
        {
            string jpgFileName = $@"{Path.GetDirectoryName(pngFileName)}\{Path.GetFileNameWithoutExtension(pngFileName)}.jpg";
            PngFileToJpgFile(pngFileName, jpgFileName);
        }
        #endregion

        #region JPG to PDF
        public static void JpgFileToPdfFile(string jpgFileName, string pdfFileName)
        {
            using (PdfDocument document = new PdfDocument())
            {
                for (int i = 0; i < 5; i++)
                {
                    PdfPage page = document.AddPage();

                    XGraphics gfx = XGraphics.FromPdfPage(page);
                    XImage image = XImage.FromFile(jpgFileName);

                    page.Height = XUnit.FromPoint(image.PointHeight);
                    page.Width = XUnit.FromPoint(image.PointWidth);

                    gfx.DrawImage(image, 0, 0);
                }

                document.Save(pdfFileName);
            }
        }
        #endregion

        #region PDF to JPG
        // Только если вся страница - картинка
        public static void PfgFileToJpgFile(string pdfFileName, string jpgFolderName)
        {
            PdfDocument document = PdfReader.Open(pdfFileName);
            var newDir = Directory.CreateDirectory(jpgFolderName);
            var path = newDir.FullName;
                        
            
            for(int i = 0; i < document.PageCount; i++)
            {
                var page = document.Pages[i];
                var resources = page.Elements.GetDictionary("/Resources");
                var xObjects = resources.Elements.GetDictionary("/XObject");
                var items = xObjects.Elements.Values;
                foreach(var item in items)
                {
                    PdfReference reference = item as PdfReference;
                    PdfDictionary xObject = reference.Value as PdfDictionary;

                    byte[] stream = xObject.Stream.Value;

                    var imgFile = File.Create(Path.Combine(path, $"page{i + 1}.jpg"));
                    using (BinaryWriter bw = new BinaryWriter(imgFile))
                    {
                        bw.Write(stream);
                    }
                }
            }
            
        }
        #endregion
    }
}
