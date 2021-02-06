namespace FilepathCheckerWPF
{
    /// <summary>
    /// A wrapper class that represents a filepath
    /// and it's existing-status.
    /// </summary>
    public class FileWrapper : IFileWrapper
    {
        public bool FileExists { get; set; }
        public string Filepath { get; set; }    
    }
}
