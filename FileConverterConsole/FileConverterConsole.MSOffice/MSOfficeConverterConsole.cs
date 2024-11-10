using CommandLine;
using FileConverterLib.MSOffice;

namespace FileConverterConsole.MSOffice
{
    internal class MSOfficeConverterConsole
    {
        static void Main(string[] args)
        {
            var cmdOptions = Parser.Default.ParseArguments<CommandLineOptions>(args).Value;

            switch (cmdOptions.Method)
            {
                case "DocxFileToPdfFile":
                    if (cmdOptions.pdfFileName == null)
                        MSOfficeConverter.DocxFileToPdfFile(cmdOptions.wordFileName);
                    else
                        MSOfficeConverter.DocxFileToPdfFile(cmdOptions.wordFileName, cmdOptions.pdfFileName);
                    break;

                case "PdfFileToDocxFile":
                    if (cmdOptions.wordFileName == null)
                        MSOfficeConverter.PdfFileToDocxFile(cmdOptions.pdfFileName);
                    else
                        MSOfficeConverter.PdfFileToDocxFile(cmdOptions.pdfFileName, cmdOptions.wordFileName);
                    break;

                case "PptxFileToPdfFile":
                    if (cmdOptions.pdfFileName == null)
                        MSOfficeConverter.PptxFileToPdfFile(cmdOptions.pptxFileName);
                    else
                        MSOfficeConverter.PptxFileToPdfFile(cmdOptions.pptxFileName, cmdOptions.pdfFileName);
                    break;

            }
        }
    }
}
