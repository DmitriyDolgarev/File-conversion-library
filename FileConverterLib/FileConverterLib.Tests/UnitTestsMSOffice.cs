using FileConverterLib.MSOffice;

namespace FileConverterLib.Tests
{
    [TestClass]
    public class UnitTestsMSOffice : AbstractUnitTests
    {
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
            File.Copy(Path.Combine(pathTestfiles, $"{filename}.docx"), Path.Combine(pathResult, $"{filename}.docx"));
            
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
            File.Copy(Path.Combine(pathTestfiles, $"{filename}.pptx"), Path.Combine(pathResult, $"{filename}.pptx"));

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
            File.Copy(Path.Combine(pathTestfiles, $"{filename}.pdf"), Path.Combine(pathResult, $"{filename}.pdf"));

            MSOfficeConverter.PdfFileToDocxFile(Path.Combine(pathResult, filename));

            Assert.IsTrue(File.Exists(Path.Combine(pathResult, $"{filename}.docx")));

        }
        #endregion
    }
}