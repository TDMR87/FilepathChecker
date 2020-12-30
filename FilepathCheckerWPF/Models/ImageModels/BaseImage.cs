using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FilepathCheckerWPF
{
    public abstract class BaseImage
    {
        /// <summary>
        /// Returns the filepath to the image resource.
        /// </summary>
        /// <returns></returns>
        public virtual string Path()
        {
            throw new NotImplementedException();
        }
    }
}
