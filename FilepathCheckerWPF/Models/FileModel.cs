using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace FilepathCheckerWPF
{
    public class FileModel : IFileModel
    {
        public bool FileExists { get; set; } = false;
        public string Filepath { get; set; } = "Filepath not set.";
    }
}
