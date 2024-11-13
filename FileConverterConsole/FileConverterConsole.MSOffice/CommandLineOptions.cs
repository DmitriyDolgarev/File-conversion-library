using CommandLine;

namespace FileConverterConsole.MSOffice
{
    internal class CommandLineOptions
    {
        [Option(longName: "method", Required = true)]
        public string? Method { get; set; }

        [Option(longName: "wordFileName", Required = false, Default = null)]
        public string? wordFileName { get; set; }

        [Option(longName: "pdfFileName", Required = false, Default = null)]
        public string? pdfFileName { get; set; }

        [Option(longName: "pptxFileName", Required = false, Default = null)]
        public string? pptxFileName { get; set; }
    }
}
