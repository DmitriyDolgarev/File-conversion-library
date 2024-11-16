using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System.IO.Compression;

using FileConverterLib.Images;
using FileConverterLib.PDF;
using FileConverterLib.MSOffice;
using FileConverterLib.LibreOffice;
using System.Net;

namespace FileConverterLib.Tests
{
    [TestClass]
    public class UnitTestsImages : AbstractUnitTests
    {
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
            File.Copy(Path.Combine(pathTestfiles, $"{filename}.png"), Path.Combine(pathResult, $"{filename}.png"));
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
            File.Copy(Path.Combine(pathTestfiles, $"{filename}.jpg"), Path.Combine(pathResult, $"{filename}.jpg"));
            ImageConverter.JpgFileToPngFile(Path.Combine(pathResult, filename));

            //проверяем сущестсвует ли файл с таким путём
            Assert.IsTrue(File.Exists(Path.Combine(pathResult, $"{filename}.png")));

        }

        #endregion
    }
}