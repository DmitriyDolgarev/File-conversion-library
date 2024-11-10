using CommandLine;

namespace FileConverterConsole.PDF
{
    internal class CommandLineOptions
    {
        [Option(longName: "method", Required = true)]
        public string? Method { get; set; }

        // MergePDFs
        [Option(longName: "pdfFiles", Required = false, Default = null)]
        public IEnumerable<string>? pdfFiles { get; set; }

        [Option(longName: "pdfOutput", Required = false, Default = null)]
        public string? pdfOutput { get; set; }

        // SplitPDF
        [Option(longName: "pdfInput", Required = false, Default = null)]
        public string? pdfInput { get; set; }

        [Option(longName: "pageSplitFrom", Required = false, Default = null)]
        public int? pageSplitFrom { get; set; }

        [Option(longName: "pdf1Output", Required = false, Default = null)]
        public string? pdf1Output { get; set; }

        [Option(longName: "pdf2Output", Required = false, Default = null)]
        public string? pdf2Output { get; set; }

        // JpgFilesToPdfFile
        [Option(longName: "jpgFiles", Required = false, Default = null)]
        public IEnumerable<string>? jpgFiles { get; set; }

        // JpgFilesToPdfFile and PdfFileToJpgFiles
        [Option(longName: "pdfFileName", Required = false, Default = null)]
        public string? pdfFileName { get; set; }

        // PdfFileToJpgFiles
        [Option(longName: "jpgFolderName", Required = false, Default = null)]
        public string? jpgFolderName { get; set; }

        [Option(longName: "zip", Required = false, Default = null)]
        public bool? zip { get; set; }
    }
}
