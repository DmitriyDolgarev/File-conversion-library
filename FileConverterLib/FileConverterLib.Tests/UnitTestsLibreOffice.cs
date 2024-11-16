using FileConverterLib.LibreOffice;

namespace FileConverterLib.Tests
{
    [TestClass]
    public class UnitTestsLibreOffice : AbstractUnitTests
    {
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
            File.Copy(Path.Combine(pathTestfiles, $"{filename}.pdf"), Path.Combine(pathResult, $"{filename}.pdf"));

            LibreOfficeConverter.PdfFileToDocxFile(Path.Combine(pathResult, filename));

            Assert.IsTrue(File.Exists(Path.Combine(pathResult, $"{filename}.docx")));

        }
        #endregion

        #region Test_WordToPdf_LibreOffice
        [TestMethod]
        public void Test_WordToPdf_LibreOffice_2params()
        {
            string filename = "test_word";
            LibreOfficeConverter.DocxFileToPdfFile(Path.Combine(pathTestfiles, filename), Path.Combine(pathResult, filename));

            Assert.IsTrue(File.Exists(Path.Combine(pathResult, $"{filename}.pdf")));

        }

        [TestMethod]
        public void Test_WordToPdf_LibreOffice_1param()
        {
            string filename = "test_word";
            File.Copy(Path.Combine(pathTestfiles, $"{filename}.docx"), Path.Combine(pathResult, $"{filename}.docx"));

            LibreOfficeConverter.DocxFileToPdfFile(Path.Combine(pathResult, $"{filename}_conv.pdf"));

            Assert.IsTrue(File.Exists(Path.Combine(pathResult, $"{filename}_conv.pdf")));

        }
        #endregion

        #region Test_PptxToPdf_LibreOffice
        [TestMethod]
        public void Test_PptxToPdf_LibreOffice_2params()
        {
            string filename = "test_powerpoint";
            LibreOfficeConverter.PptxFileToPdfFile(Path.Combine(pathTestfiles, filename), Path.Combine(pathResult, filename));

            Assert.IsTrue(File.Exists(Path.Combine(pathResult, $"{filename}.pdf")));

        }

        [TestMethod]
        public void Test_PptxToPdf_LibreOffice_1param()
        {
            string filename = "test_powerpoint";
            File.Copy(Path.Combine(pathTestfiles, $"{filename}.pptx"), Path.Combine(pathResult, $"{filename}.pptx"));

            LibreOfficeConverter.PptxFileToPdfFile(Path.Combine(pathResult, $"{filename}_conv.pdf"));

            Assert.IsTrue(File.Exists(Path.Combine(pathResult, $"{filename}_conv.pdf")));

        }
        #endregion
    }
}