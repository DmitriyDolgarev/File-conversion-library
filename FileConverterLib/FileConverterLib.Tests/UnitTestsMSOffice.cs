#if WINDOWS
using FileConverterLib.MSOffice;

namespace FileConverterLib.Tests
{
    [TestClass]
    public class UnitTestsMSOffice : AbstractUnitTests
    {
        [ClassInitialize]
        public static void BeforeTests_MSOffice(TestContext context)
        {
            var filesToCopy = new string[] { "test_word.docx", "test_powerpoint.pptx", "test_pdf_2.pdf" };

            foreach (var file in filesToCopy)
                File.Copy(Path.Combine(pathTestfiles, file), Path.Combine(pathResult, file));

        }

        #region Test_WordToPdf_MSOffice
        [TestMethod]
        public void Test_WordToPdf_MSOffice_2params()
        {
            string filename = "test_word";
            MSOfficeConverter.DocxFileToPdfFile(Path.Combine(pathTestfiles, filename), Path.Combine(pathResult, filename));
            
            Assert.IsTrue(File.Exists(Path.Combine(pathResult, $"{filename}.pdf")));

        }

        [TestMethod]
        public void Test_WordToPdf_MSOffice_1param()
        {
            string filename = "test_word";
            
            MSOfficeConverter.DocxFileToPdfFile(Path.Combine(pathResult, filename));
            
            Assert.IsTrue(File.Exists(Path.Combine(pathResult, $"{filename}.pdf")));

        }
        #endregion

        #region Test_PptxToPdf_MSOffice
        [TestMethod]
        public void Test_PptxToPdf_MSOffice_2params()
        {
            string filename = "test_powerpoint";
            MSOfficeConverter.PptxFileToPdfFile(Path.Combine(pathTestfiles, filename), Path.Combine(pathResult, filename));

            Assert.IsTrue(File.Exists(Path.Combine(pathResult, $"{filename}.pdf")));

        }

        [TestMethod]
        public void Test_PptxToPdf_MSOffice_1param()
        {
            string filename = "test_powerpoint";

            MSOfficeConverter.PptxFileToPdfFile(Path.Combine(pathResult, filename));

            Assert.IsTrue(File.Exists(Path.Combine(pathResult, $"{filename}.pdf")));

        }
        #endregion

        #region Test_PdfToWord_MSOffice
        [TestMethod]
        public void Test_PdfToWord_MSOffice_2params()
        {
            string filename = "test_pdf_2";
            MSOfficeConverter.PdfFileToDocxFile(Path.Combine(pathTestfiles, filename), Path.Combine(pathResult, filename));

            Assert.IsTrue(File.Exists(Path.Combine(pathResult, $"{filename}.docx")));

        }

        [TestMethod]
        public void Test_PdfToWord_MSOficce_1param()
        {
            string filename = "test_pdf_2";

            MSOfficeConverter.PdfFileToDocxFile(Path.Combine(pathResult, filename));

            Assert.IsTrue(File.Exists(Path.Combine(pathResult, $"{filename}.docx")));

        }
        #endregion
    }
}
#endif