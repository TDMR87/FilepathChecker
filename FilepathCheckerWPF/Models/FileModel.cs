using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace FilepathCheckerWPF.Models
{
    public class FileModel
    {
        public bool FileExists { get; set; } = false;
        public string Filepath { get; set; }
    }
}
