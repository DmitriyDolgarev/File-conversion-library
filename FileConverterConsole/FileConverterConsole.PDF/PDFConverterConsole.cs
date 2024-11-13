using CommandLine;
using FileConverterLib.PDF;
using System.Linq;

namespace FileConverterConsole.PDF
{
    internal class PDFConverterConsole
    {
        static void Main(string[] args)
        {
            var cmdOptions = Parser.Default.ParseArguments<CommandLineOptions>(args).Value;

            switch (cmdOptions.Method)
            {
                case "MergePDFs":

                    string[] pdfFiles = cmdOptions.pdfFiles.ToArray();
                    PDFConverter.MergePDFs(pdfFiles, cmdOptions.pdfOutput);
                    break;

                case "SplitPDF":
                    if (cmdOptions.pdf1Output == null && cmdOptions.pdf2Output == null)
                        PDFConverter.SplitPDF(cmdOptions.pdfInput, cmdOptions.pageSplitFrom.Value);
                    else
                        PDFConverter.SplitPDF(cmdOptions.pdfInput, cmdOptions.pageSplitFrom.Value, cmdOptions.pdf1Output, cmdOptions.pdf2Output);
                    break;

                case "JpgFilesToPdfFile":
                    string[] jpgFiles = cmdOptions.jpgFiles.ToArray();
                    PDFConverter.JpgFilesToPdfFile(jpgFiles, cmdOptions.pdfFileName);
                    break;

                case "PdfFileToJpgFiles":
                    bool zip;
                    if (cmdOptions.zip == null)
                        zip = false;
                    else
                        zip = cmdOptions.zip.Value;

                    if (cmdOptions.jpgFolderName == null)
                        PDFConverter.PdfFileToJpgFiles(cmdOptions.pdfFileName, zip);
                    else
                        PDFConverter.PdfFileToJpgFiles(cmdOptions.pdfInput, cmdOptions.jpgFolderName, zip);
                    break;
            }
        }
    }
}
