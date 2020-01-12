using System.Threading.Tasks;

namespace FilepathCheckerWPF
{
    public interface ILogger
    {
        void Close();
        void Dispose();
        void LogFileNotFound(string line);
        Task LogFileNotFoundAsync(string line);
    }
}