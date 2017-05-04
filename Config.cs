using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Utility
{
    public class Config
    {
        public List<Task> Tasks { get; private set; }

        public Config(IEnumerable<Task> tasks)
        {
            Tasks = new List<Task>();
            Tasks.AddRange(tasks);
        }
    }
}