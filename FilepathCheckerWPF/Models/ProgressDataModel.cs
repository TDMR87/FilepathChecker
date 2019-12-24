using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FilepathCheckerWPF.Models
{
    public class ProgressReportModel
    {
        public List<FileModel> FilesChecked { get; set; } = new List<FileModel>();
        public int PercentageCompleted { get; set; } = 0;
    }

    public class ProgressReportModelV2
    {
        public List<string> Filepaths { get; set; } = new List<string>();
        public int PercentageCompleted { get; set; } = 0;
    }
}
