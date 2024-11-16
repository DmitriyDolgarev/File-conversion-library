namespace FileConverterLib.Utils
{
    public class FileConverterUtils
    {
        public static string GetFileNameInSameFolder(string fileName)
        {
            var dirPath = Path.GetDirectoryName(fileName);
            var _fileName = Path.GetFileNameWithoutExtension(fileName);

            return Path.Combine(dirPath, _fileName);
        }

        public static string GetFileNameInSameFolder(string fileName, string extension)
        {
            return GetFileNameInSameFolder(fileName) + "." + extension;
        }

        public static string GetCorrectedPath(string fileName, string extension)
        {
            if (Path.GetExtension(fileName) == "")
                fileName = Path.ChangeExtension(fileName, extension);

            return Path.GetFullPath(fileName);
        }
    }
}
