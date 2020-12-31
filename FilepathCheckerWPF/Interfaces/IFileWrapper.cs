namespace FilepathCheckerWPF
{
    public interface IFileWrapper
    {
        bool FileExists { get; set; }
        string Filepath { get; set; }
    }
}