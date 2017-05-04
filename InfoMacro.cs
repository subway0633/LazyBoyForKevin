using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Utility
{
    public class InfoMacro
    {
        public string Command { get; private set; }
        public string CheckCell { get; private set; }

        public string CheckValue { get; private set; }

        public InfoMacro(string command, string checkCell, string checkValue)
        {
            Command = command;
            CheckCell = checkCell.ToUpper();
            CheckValue = checkValue;
        }
    }
}