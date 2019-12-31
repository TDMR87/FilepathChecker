using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FilepathCheckerWPF
{
    public class ErrorMessage : IMessage
    {
        public string Content { get; set; }
    }
}
