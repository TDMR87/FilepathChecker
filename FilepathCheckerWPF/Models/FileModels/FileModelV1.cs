
namespace FilepathCheckerWPF
{
    public class FileModelV1 : IFileModel
    {
        public bool FileExists { get; set; } = false;
        public string Filepath { get; set; } = "Filepath not set.";     
    }
}
