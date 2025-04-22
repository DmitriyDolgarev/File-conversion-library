namespace FileConverterLib.LibreOffice
{
    public partial class LibreOfficeConverter
    {
        #region Word to PDF
        public static async Task DocxFileToPdfFileAsync(string wordFileName, string pdfFileName) => await Task.Run(() => DocxFileToPdfFile(wordFileName, pdfFileName));
        public static async Task DocxFileToPdfFileAsync(string wordFileName) => await Task.Run(() => DocxFileToPdfFile(wordFileName));
        #endregion

        #region PDF to Word
        public static async Task PdfFileToDocxFileAsync(string pdfFileName, string wordFileName) => await Task.Run(() => PdfFileToDocxFile(pdfFileName, wordFileName));
        public static async Task PdfFileToDocxFileAsync(string pdfFileName) => await Task.Run(() => PdfFileToDocxFile(pdfFileName));
        #endregion

        #region Pptx to PDF
        public static async Task PptxFileToPdfFileAsync(string pptxFileName, string pdfFileName) => await Task.Run(() => PptxFileToPdfFile(pptxFileName, pdfFileName));
        public static async Task PptxFileToPdfFileAsync(string pptxFileName) => await Task.Run(() => PptxFileToPdfFile(pptxFileName));
        #endregion
    }
}
