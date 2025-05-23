﻿using CommandLine;
using System.Collections.Generic;

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

        [Option(longName: "splitString", Required = false, Default = null)]
        public string? splitString { get; set; }

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
