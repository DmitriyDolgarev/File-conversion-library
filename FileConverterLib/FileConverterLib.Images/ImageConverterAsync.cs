namespace FileConverterLib.Images
{
    public partial class ImageConverter
    {
        #region JPG to PNG
        public static async Task JpgFileToPngFileAsync(string jpgFileName, string pngFileName) => await Task.Run(() => JpgFileToPngFile(jpgFileName, pngFileName));
        public static async Task JpgFileToPngFileAsync(string jpgFileName) => await Task.Run(() => JpgFileToPngFile(jpgFileName));
        public static async Task<byte[]> JpgBytesToPngBytesAsync(byte[] jpgBytes) => await Task.Run(() => JpgBytesToPngBytes(jpgBytes));
        #endregion

        #region PNG TO JPG
        public static async Task PngFileToJpgFileAsync(string pngFileName, string jpgFileName) => await Task.Run(() => PngFileToJpgFile(pngFileName, jpgFileName));
        public static async Task PngFileToJpgFileAsync(string pngFileName) => await Task.Run(() => PngFileToJpgFile(pngFileName));
        public static async Task<byte[]> PngBytesToJpgBytesAsync(byte[] pngBytes) => await Task.Run(() => PngBytesToJpgBytes(pngBytes));
        #endregion
    }
}