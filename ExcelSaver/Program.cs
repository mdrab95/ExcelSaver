//using System;
public interface IExcelService { }

namespace ExcelSaver
{
    class Program
    {
        static void Main(string[] args)
        {
            /* How does this work:
             * 1. Program gets workbooks from active Excel Application
             *    If there is more than 1 "EXCEL" process, program may work incorrect.
             *    Next, it checks if directory 'C:\ExcelSaver' exists. If not - it creates it.
             *    After it, it saves all opened workbooks in this directory.
             *    If file with same name exist, program overwrites it.
             * 2. Program releases used memory (FinalReleaseComObject(workbook)).
             * 3. Program opens all saved files. 
             * 4. Program again releases used memory. 
             * */
            ExcelService esrv = new ExcelService();
            Operations oper = new Operations();
            esrv.SaveOpenedFiles();
            oper.killExcel();
            esrv.OpenAllFiles();
            //Console.ReadKey();
        }
    }
}
