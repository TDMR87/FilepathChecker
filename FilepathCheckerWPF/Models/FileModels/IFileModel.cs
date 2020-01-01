namespace FilepathCheckerWPF
{
    public interface IFileModel
    {
        bool FileExists { get; set; }
        string Filepath { get; set; }
    }
}