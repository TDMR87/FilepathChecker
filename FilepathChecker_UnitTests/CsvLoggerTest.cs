using FilepathCheckerWPF;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;

namespace FilepathChecker_UnitTests
{
    [TestClass]
    public class CsvLoggerTest
    {
        private string testString = "Test;Test";

        [TestMethod]
        public void Write()
        {
            // Arrange
            CsvLogger logger = new CsvLogger();
            string logPath = CsvLogger.GetPath();

            // Act
            logger.Write(testString);
            logger.Close();
            logger.Dispose();

            // Assert
            if (!File.Exists(logPath) || string.IsNullOrWhiteSpace(File.ReadAllText(logPath)))
            {
                Assert.Fail();
            }

            // Delete files
            if (File.Exists(logPath))
            {
                File.Delete(logPath);
            }
        }
    }
}
