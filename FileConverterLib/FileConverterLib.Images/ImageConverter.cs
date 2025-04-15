using SkiaSharp;
using FileConverterLib.Utils;
using System.Diagnostics;

namespace FileConverterLib.Images
{
    public class ImageConverter
    {
        #region JPG to PNG
        public static void JpgFileToPngFile(string jpgFileName, string pngFileName)
        {
            jpgFileName = FileConverterUtils.GetCorrectedPath(jpgFileName, "jpg");
            pngFileName = FileConverterUtils.GetCorrectedPath(pngFileName, "png");

            using (var img = SKImage.FromEncodedData(jpgFileName))
            {
                using (var data = img.Encode(SKEncodedImageFormat.Png, 100))
                {
                    using (var stream = File.OpenWrite(pngFileName))
                    {
                        data.SaveTo(stream);
                    }
                }
            }
        }
        public static void JpgFileToPngFile(string jpgFileName)
        {
            JpgFileToPngFile(jpgFileName, FileConverterUtils.GetFileNameInSameFolder(jpgFileName));
        }
        public static byte[] JpgBytesToPngBytes(byte[] jpgBytes)
        {
            using (var img = SKImage.FromEncodedData(jpgBytes))
            {
                using (var data = img.Encode(SKEncodedImageFormat.Png, 100))
                {
                    var pngBytes = data.ToArray();
                    return pngBytes;
                }
            }
        }
        #endregion

        #region PNG TO JPG
        public static void PngFileToJpgFile(string pngFileName, string jpgFileName)
        {
            pngFileName = FileConverterUtils.GetCorrectedPath(pngFileName, "png");
            jpgFileName = FileConverterUtils.GetCorrectedPath(jpgFileName, "jpg");

            using (var img = SKImage.FromEncodedData(pngFileName))
            {
                using (var data = img.Encode(SKEncodedImageFormat.Jpeg, 100))
                {
                    using (var stream = File.OpenWrite(jpgFileName))
                    {
                        data.SaveTo(stream);
                    }
                }
            }
        }
        public static void PngFileToJpgFile(string pngFileName)
        {
            PngFileToJpgFile(pngFileName, FileConverterUtils.GetFileNameInSameFolder(pngFileName));
        }
        public static byte[] PngBytesToJpgBytes(byte[] pngBytes)
        {
            using (var img = SKImage.FromEncodedData(pngBytes))
            {
                using (var data = img.Encode(SKEncodedImageFormat.Jpeg, 100))
                {
                    var jpgBytes = data.ToArray();
                    return jpgBytes;
                }
            }
        }
        #endregion
    }
}