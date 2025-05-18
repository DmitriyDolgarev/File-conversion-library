using FileConverterLib.LibreOffice;

namespace FileConverterLib.Tests
{
    [TestClass]
    public class UnitTestsLibreOffice : AbstractUnitTests
    {
        [ClassInitialize]
        public static void BeforeTests_LibreOffice(TestContext context)
        {
            var filesToCopy = new string[] { "test_pdf_1.pdf", "test_word.docx", "test_powerpoint.pptx" };

            foreach (var file in filesToCopy)
                File.Copy(Path.Combine(pathTestfiles, file), Path.Combine(pathResult, file));

        }

        #region Test_PdfToWord_LibreOffice
        [TestMethod]
        public void Test_PdfToWord_LibreOffice_2params()
        {
            string filename = "test_pdf_1";
            LibreOfficeConverter.PdfFileToDocxFile(Path.Combine(pathTestfiles, filename), Path.Combine(pathResult, $"{filename}_conv.docx"));

            Assert.IsTrue(File.Exists(Path.Combine(pathResult, $"{filename}_conv.docx")));

        }

        [TestMethod]
        public void Test_PdfToWord_LibreOffice_1param()
        {
            string filename = "test_pdf_1";

            LibreOfficeConverter.PdfFileToDocxFile(Path.Combine(pathResult, filename));

            Assert.IsTrue(File.Exists(Path.Combine(pathResult, $"{filename}.docx")));

        }
        #endregion

        #region Test_WordToPdf_LibreOffice
        [TestMethod]
        public void Test_WordToPdf_LibreOffice_2params()
        {
            string filename = "test_word";
            LibreOfficeConverter.DocxFileToPdfFile(Path.Combine(pathTestfiles, filename), Path.Combine(pathResult, $"{filename}_conv.pdf"));

            Assert.IsTrue(File.Exists(Path.Combine(pathResult, $"{filename}_conv.pdf")));

        }

        [TestMethod]
        public void Test_WordToPdf_LibreOffice_1param()
        {
            string filename = "test_word";

            LibreOfficeConverter.DocxFileToPdfFile(Path.Combine(pathResult, filename));

            Assert.IsTrue(File.Exists(Path.Combine(pathResult, $"{filename}.pdf")));

        }
        #endregion

        #region Test_PptxToPdf_LibreOffice
        [TestMethod]
        public void Test_PptxToPdf_LibreOffice_2params()
        {
            string filename = "test_powerpoint";
            LibreOfficeConverter.PptxFileToPdfFile(Path.Combine(pathTestfiles, filename), Path.Combine(pathResult, $"{filename}_conv.pdf"));

            Assert.IsTrue(File.Exists(Path.Combine(pathResult, $"{filename}_conv.pdf")));

        }

        [TestMethod]
        public void Test_PptxToPdf_LibreOffice_1param()
        {
            string filename = "test_powerpoint";

            LibreOfficeConverter.PptxFileToPdfFile(Path.Combine(pathResult, filename));

            Assert.IsTrue(File.Exists(Path.Combine(pathResult, $"{filename}.pdf")));

        }
        #endregion
    }
}