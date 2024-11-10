using CommandLine;

namespace FileConverterConsole.Images
{
    internal class CommandLineOptions
    {
        [Option(longName: "method", Required = true)]
        public string? Method { get; set; }

        [Option(longName: "pngFileName", Required = false, Default = null)]
        public string? PngFileName { get; set; }

        [Option(longName: "jpgFileName", Required = false, Default = null)]
        public string? JpgFileName { get; set; }
    }
}
