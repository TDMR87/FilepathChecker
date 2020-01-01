
namespace FilepathCheckerWPF
{
    public class FileModel : IFileModel
    {
        public bool FileExists { get; set; } = false;
        public string Filepath { get; set; } = "Filepath not set.";     
    }
}
