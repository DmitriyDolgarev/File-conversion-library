namespace FileConverterLib.PDF
{
    public partial class PdfConverter
    {
        #region Merge PDFs
        public static async Task MergePdfFilesAsync(string[] pdfFiles, string pdfOutput) => await Task.Run(() => MergePdfFiles(pdfFiles, pdfOutput));
        public static async Task<byte[]> MergePdfBytesAsync(IEnumerable<byte[]> pdfFiles) => await Task.Run(() => MergePdfBytes(pdfFiles));
        #endregion

        #region Split PDF
        public static async Task SplitPdfFileAsync(string pdfInput, int pageSplitFrom, string pdf1Output, string pdf2Output) => await Task.Run(() => SplitPdfFile(pdfInput, pageSplitFrom, pdf1Output, pdf2Output));
        public static async Task SplitPdfFileAsync(string pdfInput, int pageSplitFrom) => await Task.Run(() => SplitPdfFile(pdfInput, pageSplitFrom));
        #endregion

        #region JPG to PDF
        public static async Task JpgFilesToPdfFileAsync(string[] jpgFiles, string pdfFileName) => await Task.Run(() => JpgFilesToPdfFile(jpgFiles, pdfFileName));
        public static async Task<byte[]> JpgBytesToPdfBytesAsync(IEnumerable<byte[]> jpgBytes) => await Task.Run(() => JpgBytesToPdfBytes(jpgBytes));
        #endregion

        #region PDF to JPG
        public static async Task PdfFileToJpgFilesAsync(string pdfFileName, string jpgFolderName, bool zip = false) => await Task.Run(() => PdfFileToJpgFiles(pdfFileName, jpgFolderName, zip));
        public static async Task PdfFileToJpgFilesAsync(string pdfFileName, bool zip = false) => await Task.Run(() => PdfFileToJpgFiles(pdfFileName, zip));
        public static async Task<IEnumerable<byte[]>> PdfBytesToJpgBytesAsync(byte[] pdfBytes) => await Task.Run(() => PdfBytesToJpgBytes(pdfBytes));
        public static async Task<byte[]> PdfBytesToJpgBytesZipAsync(byte[] pdfBytes) => await Task.Run(() => PdfBytesToJpgBytesZip(pdfBytes));            
        #endregion
    }
}
