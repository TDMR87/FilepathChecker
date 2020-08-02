using System;
using System.Globalization;
using System.IO;
using System.Threading.Tasks;

namespace FilepathCheckerWPF
{
    // Use sealed so you can use the "easier" dispose pattern, if it is not sealed
    // you should create a `protected virtual void Dispose(bool disposing)` method.
    public sealed class CsvLogger : IDisposable, ILogger
    {
        private StreamWriter _writer;
        private static string _fileDirectory = "";
        private static string _fileName = "";
        private static string _fileExtension = "";
        private readonly string _titleRow = "";

        /// <summary>
        /// Creates an instance of a CSV logger. 
        /// Instanciating this class creates a log file to the application root folder 
        /// and writes one title line to that file.
        /// </summary>
        public CsvLogger()
        {
            _fileDirectory = AppDomain.CurrentDomain.BaseDirectory; // The application root directory
            _fileName = $"ERRORS {DateTime.Now.ToString("MM-dd-yyyy HH-mm-ss", CultureInfo.InvariantCulture)}";
            _fileExtension = ".csv";
            _titleRow = "Error;Filepath";

            _writer = new StreamWriter(Path.Combine(_fileDirectory, (_fileName + _fileExtension)));
            _writer.WriteLine(_titleRow);
        }

        /// <summary>
        /// Writes a line to the log file asynchronously.
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        public async Task WriteAsync(string text)
        {
            await _writer.WriteLineAsync($"{text}")
                .ConfigureAwait(true);

            // Not flushing here.
            // Flushing the buffer after each write makes sense in theory
            // but it made the application much slower..
        }

        /// <summary>
        /// Writes a line to the log file synchronously.
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        public void Write(string text)
        {
            _writer.WriteLine($"{text}");

            // Not flushing here.
            // Flushing the buffer after each write makes sense in theory
            // but it made the application much slower..
        }

        /// <summary>
        /// Returns the UNC path of the log file.
        /// </summary>
        /// <returns></returns>
        public static string GetPath()
        {
            return Path.Combine(_fileDirectory, _fileName + _fileExtension);
        }

        public void Close()
        {
            _writer.Close();
        }

        public void Dispose()
        {
            _writer.Dispose();
        }
    }
}
