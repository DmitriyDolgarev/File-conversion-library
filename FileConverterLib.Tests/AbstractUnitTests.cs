using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FileConverterLib.Tests
{
    [TestClass]
    public abstract class AbstractUnitTests
    {
        protected static string pathTestfiles = Path.GetFullPath(Path.Combine("..", "..", "..", "..", "FileConverterLib.Tests", "testfiles"));
        protected static string pathResult = Path.GetFullPath(Path.Combine("..", "..", "..", "..", "FileConverterLib.Tests", "resultfiles"));

        [ClassInitialize(InheritanceBehavior.BeforeEachDerivedClass)]
        public static void BeforeTests(TestContext _context)
        {         
            //создаём папку в которой будут хранится все файлы для тестов
            if (!Directory.Exists(pathResult))
                Directory.CreateDirectory(pathResult);

            // удаляем все тестовые файлы
            foreach (string f in Directory.GetFiles(pathResult))
                File.Delete(f);

            //удаляем все тестовые папки
            foreach (string d in Directory.GetDirectories(pathResult))
                Directory.Delete(d, true);
        }
    }
}
