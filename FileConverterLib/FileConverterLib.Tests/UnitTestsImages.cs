using FileConverterLib.Images;

namespace FileConverterLib.Tests
{
    [TestClass]
    public class UnitTestsImages : AbstractUnitTests
    {
        [ClassInitialize]
        public static void BeforeTests_Images(TestContext context) 
        {
            var filesToCopy = new string[] { "test_png.png", "test_jpg.jpg" };
            
            foreach (var file in filesToCopy)
                File.Copy(Path.Combine(pathTestfiles, file), Path.Combine(pathResult, file));

        }

        #region Test_PngToJpg
        [TestMethod]
        public void Test_PngToJpg_2params()
        {
            string filename = "test_png";
            ImageConverter.PngFileToJpgFile(Path.Combine(pathTestfiles, filename), Path.Combine(pathResult, filename));

            //проверяем сущестсвует ли файл с таким путём
            Assert.IsTrue(File.Exists(Path.Combine(pathResult,$"{filename}.jpg")));
        }

        [TestMethod]
        public void Test_PngToJpg_1param()
        {
            string filename = "test_png";
            ImageConverter.PngFileToJpgFile(Path.Combine(pathResult, filename));

            //проверяем сущестсвует ли файл с таким путём
            Assert.IsTrue(File.Exists(Path.Combine(pathResult,$"{filename}.jpg")));
        }

        #endregion

        #region Test_JpgToPng
        [TestMethod]
        public void Test_JpgToPng_2params()
        {
            string filename = "test_jpg";
            ImageConverter.JpgFileToPngFile(Path.Combine(pathTestfiles, filename), Path.Combine(pathResult, filename));

            //проверяем сущестсвует ли файл с таким путём
            Assert.IsTrue(File.Exists(Path.Combine(pathResult, $"{filename}.png")));
        }

        [TestMethod]
        public void Test_JpgToPng_1param()
        {
            string filename = "test_jpg";

            //копируем файл из папки testfiles в папку results, чтобы создать преобразованный файл в папке с исходным и с тем же названием
            ImageConverter.JpgFileToPngFile(Path.Combine(pathResult, filename));

            //проверяем сущестсвует ли файл с таким путём
            Assert.IsTrue(File.Exists(Path.Combine(pathResult, $"{filename}.png")));
        }

        #endregion
    }
}