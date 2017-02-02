using System;
using System.Diagnostics;

namespace ExcelSaver
{
    class Operations
    {  
        /*
        public void OpenMicrosoftExcel(string file)
        {
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.FileName = "EXCEL";
            startInfo.Arguments = file;
            Process.Start(startInfo);
        }
        */
        

        public void killExcel()
        {
            try
            {
                foreach (Process proc in Process.GetProcessesByName("EXCEL"))
                {
                    proc.Kill();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }
    }
}

