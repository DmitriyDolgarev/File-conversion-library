using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System.IO.Compression;

using FileConverterLib.PDF;

namespace FileConverterLib.Tests
{
    [TestClass]
    public class UnitTestsPdf: AbstractUnitTests
    {
        [ClassInitialize]
        public static void BeforeTests_PDF(TestContext context)
        {
            var filesToCopy = new string[] { "test_pdf_2.pdf", "test_pdf_to_word.pdf", "test_split_pdf.pdf" };

            foreach (var file in filesToCopy)
                File.Copy(Path.Combine(pathTestfiles, file), Path.Combine(pathResult, file));

        }

        #region Test_PdfToJpg
        [TestMethod]
        public void Test_PdfToJpgs_folder_2param()
        {
            string filename = "test_pdf_1";
            PdfConverter.PdfFileToJpgFiles(Path.Combine(pathTestfiles, filename), Path.Combine(pathResult, filename), false);

            //открываем pdf 
            using (PdfDocument document = PdfReader.Open(Path.Combine(pathTestfiles, $"{filename}.pdf")))
            {
                //получаем полные пути файлов из папки
                string[] files_full = Directory.GetFiles(Path.Combine(pathResult, filename));
                //проверям равно ли количество файлов в папке числу страниц в pdf
                Assert.AreEqual(document.PageCount, files_full.Length);

                //получаем названия файлов из их полных путей
                string[] files = new string[files_full.Length];
                for (int i = 0; i < files_full.Length; i++)
                {
                    files[i] = Path.GetFileName(files_full[i]);
                }

                //проверяем существуют ли файлы для каждой страницы pdf
                bool isCorrect = true;
                for (int i = 1; i <= document.PageCount; i++)
                {
                    if (Array.IndexOf(files, $"page_{i}.jpg") == -1)
                    {
                        isCorrect = false;
                        break;
                    }

                }
                Assert.IsTrue(isCorrect);
            }

        }

        [TestMethod]
        public void Test_PdfToJpgs_folder_1param()
        {
            string filename = "test_pdf_2";

            //копируем файл из папки testfiles в папку results, чтобы создать папку с изображениями в папке с исходным и с тем же названием
            PdfConverter.PdfFileToJpgFiles(Path.Combine(pathResult, filename), false);
            
            // открываем pdf
            using (PdfDocument document = PdfReader.Open(Path.Combine(pathResult, $"{filename}.pdf")))
            {
                //получаем полные пути файлов из папки
                string[] files_full = Directory.GetFiles(Path.Combine(pathResult, filename));

                //проверям равно ли количество файлов в папке числу страниц в pdf
                Assert.AreEqual(document.PageCount, files_full.Length);

                //получаем названия файлов из их полных путей
                string[] files = new string[files_full.Length];
                for (int i = 0; i < files_full.Length; i++)
                {
                    files[i] = Path.GetFileName(files_full[i]);
                }

                //проверяем существуют ли файлы для каждой страницы pdf
                bool isCorrect = true;
                for (int i = 1; i <= document.PageCount; i++)
                {
                    if (Array.IndexOf(files, $"page_{i}.jpg") == -1)
                    {
                        isCorrect = false;
                        break;
                    }

                }
                Assert.IsTrue(isCorrect);
            }
        }

        [TestMethod]
        public void Test_PdfToJpgs_zip_2param()
        {
            string filename = "test_pdf_1";
            PdfConverter.PdfFileToJpgFiles(Path.Combine(pathTestfiles, filename), Path.Combine(pathResult, filename), true);

            //проверяем существует ли архив с таким названием
            Assert.IsTrue(File.Exists(Path.Combine(pathResult, $"{filename}.zip")));

            //открываем pdf 
            using (PdfDocument document = PdfReader.Open(Path.Combine(pathTestfiles, $"{filename}.pdf")))
            {
                //открываем архив
                using (ZipArchive archive = ZipFile.OpenRead(Path.Combine(pathResult,$"{filename}.zip")))
                {
                    //получаем названия файлов из архива
                    string[] files= new string[archive.Entries.ToArray().Length];
                    for (int i=0; i < archive.Entries.ToArray().Length; i++)
                    {
                        files[i] = archive.Entries[i].FullName;
                    }

                    // проверяем равно ли количество файлов в архиве количеству страниц в pdf
                    Assert.AreEqual(document.PageCount, files.Length);

                    //проверяем существуют ли файлы для каждой страницы pdf
                    bool isCorrect = true;
                    for (int i = 1; i <= document.PageCount; i++)
                    {
                        if (Array.IndexOf(files, $"page_{i}.jpg") == -1)
                        {
                            isCorrect = false;
                            break;
                        }

                    }
                    Assert.IsTrue(isCorrect);

                }               
                
            }
            
        }

        [TestMethod]
        public void Test_PdfToJpgs_zip_1param()
        {
            string filename = "test_pdf_to_word";

            //копируем файл из папки testfiles в папку results, чтобы создать папку с изображениями в папке с исходным и с тем же названием
            PdfConverter.PdfFileToJpgFiles(Path.Combine(pathResult, filename), true);

            //проверяем существует ли архив с таким названием
            Assert.IsTrue(File.Exists(Path.Combine(pathResult, $"{filename}.zip")));

            //открываем pdf
            using (PdfDocument document = PdfReader.Open(Path.Combine(pathResult, $"{filename}.pdf")))
            {
                //открываем архив
                using (ZipArchive archive = ZipFile.OpenRead(Path.Combine(pathResult, $"{filename}.zip")))
                {
                    //получаем названия файлов из архива
                    string[] files = new string[archive.Entries.ToArray().Length];
                    for (int i = 0; i < archive.Entries.ToArray().Length; i++)
                    {
                        files[i] = archive.Entries[i].FullName;
                    }

                    // проверяем равно ли количество файлов в архиве количеству страниц в pdf
                    Assert.AreEqual(document.PageCount, files.Length);

                    //проверяем существуют ли файлы для каждой страницы pdf
                    bool isCorrect = true;
                    for (int i = 1; i <= document.PageCount; i++)
                    {
                        if (Array.IndexOf(files, $"page_{i}.jpg") == -1)
                        {
                            isCorrect = false;
                            break;
                        }

                    }
                    Assert.IsTrue(isCorrect);

                }

            }
        }

        #endregion

        #region TestJpgToPdf
        [TestMethod]
        public void Test_JpgsToPdf()
        {
            string filename = "test_jpgs_to_pdf";
            string[] files = Directory.GetFiles(Path.Combine(pathTestfiles, filename));
            PdfConverter.JpgFilesToPdfFile(files, Path.Combine(pathResult,$"{filename}.pdf"));

            using (PdfDocument document = PdfReader.Open(Path.Combine(pathResult, $"{filename}.pdf")))
            {
                Assert.AreEqual(files.Length,document.PageCount);
            }
        }
        #endregion

        #region Test_MergePdf
        [TestMethod]
        public void TestMergePdf()
        {
            string filename = "test_merge_pdf";
            string[] files = Directory.GetFiles(Path.Combine(pathTestfiles, filename));

            PdfConverter.MergePdfFiles(files, Path.Combine(pathResult, $"{filename}.pdf"));

            using (PdfDocument document = PdfReader.Open(Path.Combine(pathResult, $"{filename}.pdf")))
            {
                int pageCount = 0;
                foreach (string f in files)
                {
                    using (PdfDocument doc = PdfReader.Open(f))
                    {
                        pageCount += doc.PageCount;
                    }
                }
                Assert.AreEqual(document.PageCount, pageCount);
            }
        }
        #endregion

        #region Test_SptitPdf
        [TestMethod]
        public void TestSplitPdf_4params(){
            string filename = "test_split_pdf";
            int splitFrom = 2;

            PdfConverter.SplitPdfFile(Path.Combine(pathTestfiles, $"{filename}.pdf"), splitFrom, 
                Path.Combine(pathResult, $"{filename}_split_1.pdf"), 
                Path.Combine(pathResult, $"{filename}_split_2.pdf")
            );

            using (PdfDocument document = PdfReader.Open(Path.Combine(pathTestfiles, $"{filename}.pdf")))
            {
                using (PdfDocument doc1 = PdfReader.Open(Path.Combine(pathResult, $"{filename}_split_1.pdf")))
                {
                    Assert.AreEqual(document.PageCount - splitFrom, doc1.PageCount);
                }

                using (PdfDocument doc2 = PdfReader.Open(Path.Combine(pathResult, $"{filename}_split_2.pdf")))
                {
                    Assert.AreEqual((document.PageCount - splitFrom) + 1, doc2.PageCount) ;
                }
            }


        }

        [TestMethod]
        public void TestSplitPdf_2params()
        {
            string filename = "test_split_pdf";
            int splitFrom = 2;

            PdfConverter.SplitPdfFile(Path.Combine(pathResult, $"{filename}.pdf"), splitFrom);

            using (PdfDocument document = PdfReader.Open(Path.Combine(pathResult, $"{filename}.pdf")))
            {
                using (PdfDocument doc1 = PdfReader.Open(Path.Combine(pathResult, $"{filename}_splitted1.pdf")))
                {
                    Assert.AreEqual(document.PageCount - splitFrom, doc1.PageCount);
                }

                using (PdfDocument doc2 = PdfReader.Open(Path.Combine(pathResult, $"{filename}_splitted2.pdf")))
                {
                    Assert.AreEqual((document.PageCount - splitFrom) + 1, doc2.PageCount);
                }
            }
        }
        #endregion
    }
}