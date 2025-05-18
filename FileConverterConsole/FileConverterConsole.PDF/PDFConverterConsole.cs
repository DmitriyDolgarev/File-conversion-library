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
                    PdfConverter.MergePdfFiles(pdfFiles, cmdOptions.pdfOutput);
                    break;

                case "SplitPDF":
                    if (cmdOptions.pdfOutput == null)
                        PdfConverter.SplitPdfFile(cmdOptions.pdfInput, cmdOptions.splitString);
                    else
                        PdfConverter.SplitPdfFile(cmdOptions.pdfInput, cmdOptions.splitString, cmdOptions.pdfOutput);
                    break;

                case "JpgFilesToPdfFile":
                    string[] jpgFiles = cmdOptions.jpgFiles.ToArray();
                    PdfConverter.JpgFilesToPdfFile(jpgFiles, cmdOptions.pdfFileName);
                    break;

                case "PdfFileToJpgFiles":
                    bool zip;
                    if (cmdOptions.zip == null)
                        zip = false;
                    else
                        zip = cmdOptions.zip.Value;

                    if (cmdOptions.jpgFolderName == null)
                        PdfConverter.PdfFileToJpgFiles(cmdOptions.pdfFileName, zip);
                    else
                        PdfConverter.PdfFileToJpgFiles(cmdOptions.pdfInput, cmdOptions.jpgFolderName, zip);
                    break;
            }
        }
    }
}
