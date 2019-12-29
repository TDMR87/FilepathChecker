using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FilepathCheckerWPF
{
    // Use sealed so you can use the "easier" dispose pattern, if it is not sealed
    // you should create a `protected virtual void Dispose(bool disposing)` method.
    public sealed class CsvLogger : IDisposable
    {
        private StreamWriter writer;
        private static string _fileDirectory = "";
        private static string _fileName = "";
        private static string _fileExtension = "";
        private readonly string _titleRow = "";

        public CsvLogger()
        {
            _fileDirectory = AppDomain.CurrentDomain.BaseDirectory; // The application root directory
            _fileName = $"ERRORS {DateTime.Now.ToString("MM-dd-yyyy HH-mm-ss", CultureInfo.InvariantCulture)}";
            _fileExtension = ".csv";
            _titleRow = "Error;Filepath";

            writer = new StreamWriter(Path.Combine(_fileDirectory, (_fileName + _fileExtension)));
            writer.WriteLine(_titleRow);
        }

        /// <summary>
        /// Writes a line to the log file instance asynchronously.
        /// </summary>
        /// <param name="line"></param>
        /// <returns></returns>
        public async Task WriteLineAsync(string line)
        {
            await writer.WriteLineAsync($"File not found;{line}")
                .ConfigureAwait(true);

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
            writer.Close();
        }

        public void Dispose()
        {
            writer.Dispose();
        }

        
    }
}
