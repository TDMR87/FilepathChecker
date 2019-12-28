using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FilepathCheckerWPF
{
    // Progress model for reporting progress of iterating through IFileModel-objects
    public class ProgressReportModel
    {
        public int PercentageCompleted { get; set; } = 0;
    }

    // Progress model for reporting progress of reading text-type filepaths from excel-file
    public class ProgressReportModelV2
    {
        public List<string> Filepaths { get; set; } = new List<string>();
        public int PercentageCompleted { get; set; } = 0;
    }
}
