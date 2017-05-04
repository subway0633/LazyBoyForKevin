using SevenZip;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;

namespace Utility
{
    public class TaskTar : Task, ITask
    {
        public string TarResult { get { return _tarResult.Value(); } }
        private string _tarResult;
        private List<string> _tarSource = new List<string>();

        public TaskTar(string title, string[] source, string result) : base(title)
        {
            _tarResult = result;
            _tarSource.AddRange(source);
        }

        void ITask.Run()
        {
        }
    }
}