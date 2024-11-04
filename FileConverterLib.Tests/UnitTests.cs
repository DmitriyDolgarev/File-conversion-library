using FileConverterLib;

namespace FileConverterLib.Tests
{
    [TestClass]
    public class UnitTests
    {
        [ClassInitialize]
        public static void BeforeTests(TestContext context)
        {
            // Код выполнится до выполнения тестов
            // Надо бы почитать про контекст...
        }

        [TestMethod]
        public void TestMethod()
        {
            Assert.AreEqual(0, 0);
        }
    }
}