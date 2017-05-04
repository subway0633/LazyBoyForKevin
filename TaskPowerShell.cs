using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;

namespace Utility
{
    internal class TaskPowerShell : Task, ITask
    {
        private string _command;

        public TaskPowerShell(string title, string command) : base(title)
        {
            this._command = command;
        }

        void ITask.Run()
        {
            string command = _command.Value();
            using (PowerShell PowerShellInstance = PowerShell.Create())
            {
                try
                {
                    PowerShellInstance.AddScript(command);
                    IAsyncResult result = PowerShellInstance.BeginInvoke();

                    while (!result.IsCompleted)
                    {
                    }
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
        }
    }
}