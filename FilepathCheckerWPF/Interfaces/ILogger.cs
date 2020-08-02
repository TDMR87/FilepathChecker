using System.Threading.Tasks;

namespace FilepathCheckerWPF
{
    public interface ILogger
    {
        void Close();
        void Dispose();
        void Write(string line);
        Task WriteAsync(string line);
    }
}