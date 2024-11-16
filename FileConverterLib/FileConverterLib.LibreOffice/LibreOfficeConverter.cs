using System.Diagnostics;
using FileConverterLib.Utils;

namespace FileConverterLib.LibreOffice
{
    public class LibreOfficeConverter
    {
        // Windows: обычно @"C:\Program Files\LibreOffice\program"
        // Linux: пустая строка @""
        public static string sofficePath = @"";

        private static void sofficeLaunch(string arguments)
        {
            using (Process process = new Process())
            {
                ProcessStartInfo info = new ProcessStartInfo();
                info.FileName = "soffice";
                info.WorkingDirectory = sofficePath;
                info.Arguments = arguments;
                info.UseShellExecute = true;
                process.StartInfo = info;

                process.Start();
                process.WaitForExit();
                process.Close();
            }
        }

        #region Word to PDF
        public static void DocxFileToPdfFile(string wordFileName, string pdfFileName)
        {
            pdfFileName = FileConverterUtils.GetCorrectedPath(pdfFileName, "pdf");
            wordFileName = FileConverterUtils.GetCorrectedPath(wordFileName, "docx");

            var fileName = Path.GetFileNameWithoutExtension(wordFileName);
            var tempFolder = Directory.CreateDirectory($"~{fileName}_temp").FullName;

            var args = $"--headless --convert-to \"pdf:writer_pdf_Export\" \"{wordFileName}\" --outdir \"{tempFolder}\"";
            sofficeLaunch(args);

            File.Move(Path.Combine(tempFolder, fileName + ".pdf"), pdfFileName);
            Directory.Delete(tempFolder, true);
        }

        public static void DocxFileToPdfFile(string wordFileName)
        {
            DocxFileToPdfFile(wordFileName, FileConverterUtils.GetFileNameInSameFolder(wordFileName));
        }
        #endregion

        #region PDF to Word
        public static void PdfFileToDocxFile(string pdfFileName, string wordFileName)
        {
            pdfFileName = FileConverterUtils.GetCorrectedPath(pdfFileName, "pdf");
            wordFileName = FileConverterUtils.GetCorrectedPath(wordFileName, "docx");

            var fileName = Path.GetFileNameWithoutExtension(pdfFileName);
            var tempFolder = Directory.CreateDirectory($"~{fileName}_temp").FullName;

            var args = $"--headless --infilter=\"writer_pdf_import\" --convert-to docx \"{pdfFileName}\" --outdir \"{tempFolder}\"";
            sofficeLaunch(args);

            File.Move(Path.Combine(tempFolder, fileName + ".docx"), wordFileName);
            Directory.Delete(tempFolder, true);
        }

        public static void PdfFileToDocxFile(string pdfFileName)
        {
            PdfFileToDocxFile(pdfFileName, FileConverterUtils.GetFileNameInSameFolder(pdfFileName));
        }
        #endregion

        #region Pptx to PDF
        public static void PptxFileToPdfFile(string pptxFileName, string pdfFileName)
        {
            pdfFileName = FileConverterUtils.GetCorrectedPath(pdfFileName, "pdf");
            pptxFileName = FileConverterUtils.GetCorrectedPath(pptxFileName, "pptx");

            var fileName = Path.GetFileNameWithoutExtension(pptxFileName);
            var tempFolder = Directory.CreateDirectory($"~{fileName}_temp").FullName;

            var args = $"--headless --convert-to \"pdf\" \"{pptxFileName}\" --outdir \"{tempFolder}\"";
            sofficeLaunch(args);

            File.Move(Path.Combine(tempFolder, fileName + ".pdf"), pdfFileName);
            Directory.Delete(tempFolder, true);
        }

        public static void PptxFileToPdfFile(string pptxFileName)
        {
            PptxFileToPdfFile(pptxFileName, FileConverterUtils.GetFileNameInSameFolder(pptxFileName));
        }
        #endregion
    }
}
