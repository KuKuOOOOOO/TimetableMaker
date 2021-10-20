using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TimetableMaker
{
    class ShutdownProcess
    {
        public void ExcelProcess()
        {
            var processes = from p in Process.GetProcessesByName("EXCEL") select p;
            foreach (var process in processes)
            {
                if (process.MainWindowTitle == "")
                {
                    process.Kill();
                    process.WaitForExit();
                    process.Dispose();
                }
            }
        }
    }
}
