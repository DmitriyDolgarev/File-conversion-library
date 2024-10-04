using System.Drawing;
using System.Drawing.Imaging;

namespace FileConverterLib
{
    public class FileConverter
    {
        #region JPG to PNG
        public static void JpgFileToPngFile(string jpgFileName, string pngFileName)
        {
            if (Path.GetExtension(jpgFileName) == "")
                jpgFileName += ".jpg";
            if (Path.GetExtension(pngFileName) == "")
                pngFileName += ".png";

            using (var img = Image.FromFile(jpgFileName))
            {
                img.Save(pngFileName, ImageFormat.Png);
            }
        }

        public static void JpgFileToPngFile(string jpgFileName)
        {
            string pngFileName = $@"{Path.GetDirectoryName(jpgFileName)}\{Path.GetFileNameWithoutExtension(jpgFileName)}.png";
            JpgFileToPngFile(jpgFileName, pngFileName);
        }
        #endregion

        #region PNG TO JPG
        public static void PngFileToJpgFile(string pngFileName, string jpgFileName)
        {
            if (Path.GetExtension(pngFileName) == "")
                pngFileName += ".png";
            if (Path.GetExtension(jpgFileName) == "")
                jpgFileName += ".jpg";

            using (var img = Image.FromFile(pngFileName))
            {
                img.Save(jpgFileName, ImageFormat.Jpeg);
            }
        }

        public static void PngFileToJpgFile(string pngFileName)
        {
            string jpgFileName = $@"{Path.GetDirectoryName(pngFileName)}\{Path.GetFileNameWithoutExtension(pngFileName)}.jpg";
            PngFileToJpgFile(pngFileName, jpgFileName);
        }
        #endregion
    }
}
