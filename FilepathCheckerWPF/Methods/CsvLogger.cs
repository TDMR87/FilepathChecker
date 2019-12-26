using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FilepathCheckerWPF.Methods
{
    // Use sealed so you can use the "easier" dispose pattern, if it is not sealed
    // you should create a `protected virtual void Dispose(bool disposing)` method.
    public sealed class CsvLogger : IDisposable
    {
        private StreamWriter writer;
        private static string _filepath = AppDomain.CurrentDomain.BaseDirectory;
        private static string _filename = $"ERRORS {DateTime.Now.ToShortDateString()} {DateTime.Now.ToShortTimeString()}.csv";
        private readonly string _titleRow = "Error;Filepath";

        public CsvLogger()
        {
            _filename = _filename.Replace("/", "-");
            writer = new StreamWriter(Path.Combine(_filepath, _filename));
            writer.WriteLine(_titleRow);
        }

        public void Close()
        {
            writer.Close();
        }

        public void Dispose()
        {
            writer.Dispose();
        }

        public async Task WriteLineAsync(string line)
        {
            await writer.WriteLineAsync($"File not found;{line}")
                .ConfigureAwait(true);

            // Not flushing here.
            // Flushing the buffer after each write makes sense in theory
            // but it made the application four times slower..
        }
    }
}
