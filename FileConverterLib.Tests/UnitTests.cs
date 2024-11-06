using FileConverterLib;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using PdfSharp.Snippets.Drawing;
using System.IO.Compression;

namespace FileConverterLib.Tests
{
    [TestClass]
    public class UnitTests
    {
        private static string pathTestfiles = @"..\..\..\..\FileConverterLib.Tests\testfiles";
        private static string pathResult = @"..\..\..\..\FileConverterLib.Tests\resultfiles";

        [ClassInitialize]
        public static void BeforeTests(TestContext _context)
        {
            //создаём папку в которой будут хранится все файлы для тестов
            if (!Directory.Exists(pathResult))
                Directory.CreateDirectory(pathResult);

            // удаляем все тестовые файлы
            foreach (string f in Directory.GetFiles(pathResult))
                File.Delete(f);

            //удаляем все тестовые папки
            foreach(string d in Directory.GetDirectories(pathResult))
                Directory.Delete(d,true);
        }

        #region Test_PngToJpg

        [TestMethod]
        public void Test_PngToJpg_2params()
        {
            string filename = "test_png";
            FileConverter.PngFileToJpgFile(Path.Combine(pathTestfiles, filename), Path.Combine(pathResult, filename));

            //проверяем сущестсвует ли файл с таким путём
            Assert.IsTrue(File.Exists(Path.Combine(pathResult,$"{filename}.jpg")));

        }

        [TestMethod]
        public void Test_PngToJpg_1param()
        {
            string filename = "test_png";
            File.Copy(Path.Combine(pathTestfiles, $"{filename}.png"), Path.Combine(pathResult, $"{filename}.png"));
            FileConverter.PngFileToJpgFile(Path.Combine(pathResult, filename));

            //проверяем сущестсвует ли файл с таким путём
            Assert.IsTrue(File.Exists(Path.Combine(pathResult,$"{filename}.jpg")));

        }

        #endregion

        #region Test_JpgToPng
        [TestMethod]
        public void Test_JpgToPng_2params()
        {
            string filename = "test_jpg";
            FileConverter.JpgFileToPngFile(Path.Combine(pathTestfiles, filename), Path.Combine(pathResult, filename));

            //проверяем сущестсвует ли файл с таким путём
            Assert.IsTrue(File.Exists(Path.Combine(pathResult, $"{filename}.png")));

        }

        [TestMethod]
        public void Test_JpgToPng_1param()
        {
            string filename = "test_jpg";

            //копируем файл из папки testfiles в папку results, чтобы создать преобразованный файл в папке с исходным и с тем же названием
            File.Copy(Path.Combine(pathTestfiles, $"{filename}.jpg"), Path.Combine(pathResult, $"{filename}.jpg"));
            FileConverter.JpgFileToPngFile(Path.Combine(pathResult, filename));

            //проверяем сущестсвует ли файл с таким путём
            Assert.IsTrue(File.Exists(Path.Combine(pathResult, $"{filename}.png")));

        }

        #endregion

        #region Test_PdfToJpg
        [TestMethod]
        public void Test_PdfToJpgs_folder_2param()
        {
            string filename = "test_pdf_1";
            FileConverter.PdfFileToJpgFiles(Path.Combine(pathTestfiles, filename), Path.Combine(pathResult, filename), false);

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
            File.Copy(Path.Combine(pathTestfiles, $"{filename}.pdf"), Path.Combine(pathResult, $"{filename}.pdf"));
            FileConverter.PdfFileToJpgFiles(Path.Combine(pathResult, filename), false);
            
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
            FileConverter.PdfFileToJpgFiles(Path.Combine(pathTestfiles, filename), Path.Combine(pathResult, filename), true);

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
            File.Copy(Path.Combine(pathTestfiles, $"{filename}.pdf"), Path.Combine(pathResult, $"{filename}.pdf"));
            FileConverter.PdfFileToJpgFiles(Path.Combine(pathResult, filename), true);

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
            FileConverter.JpgFilesToPdfFile(files, Path.Combine(pathResult,$"{filename}.pdf"));

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

            FileConverter.MergePDFs(files, Path.Combine(pathResult, $"{filename}.pdf"));

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

            FileConverter.SplitPDF(Path.Combine(pathTestfiles, $"{filename}.pdf"), splitFrom, 
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

            File.Copy(Path.Combine(pathTestfiles, $"{filename}.pdf"), Path.Combine(pathResult, $"{filename}.pdf"));

            FileConverter.SplitPDF(Path.Combine(pathResult, $"{filename}.pdf"), splitFrom);

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