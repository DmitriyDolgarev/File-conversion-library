using CommandLine;
using FileConverterLib.Images;

namespace FileConverterConsole.Images
{
    internal class ImageConverterConsole
    {
        static void Main(string[] args)
        {
            var cmdOptions = Parser.Default.ParseArguments<CommandLineOptions>(args).Value;

            switch (cmdOptions.Method)
            {
                case "JpgFileToPngFile":
                    if (cmdOptions.PngFileName == null)
                        ImageConverter.JpgFileToPngFile(cmdOptions.JpgFileName);
                    else
                        ImageConverter.JpgFileToPngFile(cmdOptions.JpgFileName, cmdOptions.PngFileName);
                    break;

                case "PngFileToJpgFile":
                    if(cmdOptions.JpgFileName == null)
                        ImageConverter.PngFileToJpgFile(cmdOptions.PngFileName);
                    else
                        ImageConverter.PngFileToJpgFile(cmdOptions.PngFileName, cmdOptions.JpgFileName);
                    break;
            }
        }
    }
}
