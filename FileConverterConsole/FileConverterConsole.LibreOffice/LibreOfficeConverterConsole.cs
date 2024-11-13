using CommandLine;
using FileConverterLib.LibreOffice;

namespace FileConverterConsole.LibreOffice
{
    internal class LibreOfficeConverterConsole
    {
        static void Main(string[] args)
        {
            var cmdOptions = Parser.Default.ParseArguments<CommandLineOptions>(args).Value;

            if(cmdOptions.sofficePath != null)
                LibreOfficeConverter.sofficePath = cmdOptions.sofficePath;

            switch (cmdOptions.Method)
            {
                case "DocxFileToPdfFile":
                    if (cmdOptions.pdfFileName == null)
                        LibreOfficeConverter.DocxFileToPdfFile(cmdOptions.wordFileName);
                    else
                        LibreOfficeConverter.DocxFileToPdfFile(cmdOptions.wordFileName, cmdOptions.pdfFileName);
                    break;

                case "PdfFileToDocxFile":
                    if (cmdOptions.wordFileName == null)
                        LibreOfficeConverter.PdfFileToDocxFile(cmdOptions.pdfFileName);
                    else
                        LibreOfficeConverter.PdfFileToDocxFile(cmdOptions.pdfFileName, cmdOptions.wordFileName);
                    break;

                case "PptxFileToPdfFile":
                    if (cmdOptions.pdfFileName == null)
                        LibreOfficeConverter.PptxFileToPdfFile(cmdOptions.pptxFileName);
                    else
                        LibreOfficeConverter.PptxFileToPdfFile(cmdOptions.pptxFileName, cmdOptions.pdfFileName);
                    break;
            }
        }
    }
}
