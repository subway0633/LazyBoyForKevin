using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Utility
{
    public class InfoFile
    {
        public string FromFile { get { return source.Value(); } }
        public string Tofile { get { return target.Value(); } }

        private string source;
        private string target;

        public InfoFile SetDefaultFolder(string folder)
        {
            if (target.Length == 0) return this;
            string newTarget = target.Replace("<DEFAULT>", folder);
            return new InfoFile(source, newTarget);
        }

        public InfoFile(string source, string target)
        {
            this.source = source;
            this.target = target ?? string.Empty;
        }
    }
}