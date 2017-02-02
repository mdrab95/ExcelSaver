using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using System.Diagnostics;
using WindowsInput;
using System.Globalization;
using System.Reflection;
using System.Text;

namespace ExcelSaver
{
   public class ExcelService : IExcelService
    {
        List<string> openedFiles = new List<string>();
        Application xlApp;
        Workbooks xlBooks;
        Workbook xlBook;
        Worksheet xlSheet;

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("Oleacc.dll")]
        static extern int AccessibleObjectFromWindow(int hwnd, uint dwObjectID, byte[] riid, out ExcelWindow ptr);

        [Guid("00020893-0000-0000-C000-000000000046")]
        [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
        public interface ExcelWindow
        {
        }

        public static void BringExcelWindowToFront(Application xlApp)
        {
            SetForegroundWindow((IntPtr)xlApp.Hwnd);  // Note Hwnd is declared as int
        }


        public delegate bool EnumChildCallback(int hwnd, ref int lParam);

        [DllImport("User32.dll")]
        public static extern bool EnumChildWindows(int hWndParent, EnumChildCallback lpEnumFunc, ref int lParam);

        [DllImport("User32.dll")]
        public static extern int GetClassName(int hWnd, StringBuilder lpClassName, int nMaxCount);


        public static bool EnumChildProc(int hwndChild, ref int lParam)
        {
            StringBuilder buf = new StringBuilder(128);
            GetClassName(hwndChild, buf, 128);
            if (buf.ToString() == "EXCEL7")
            {
                lParam = hwndChild;
                return false;
            }
            return true;
        }

        /// <summary>
        /// Save all opened excel Workbooks.
        /// </summary>
        public void SaveOpenedFiles()
        {
            try
            {
                var allExcelProcesses = Process.GetProcessesByName("EXCEL");
                for (int i = 0; i < allExcelProcesses.Length; i++)
                {
                    int excelId = allExcelProcesses[i].Id;
                    int hwnd = (int)Process.GetProcessById(excelId).MainWindowHandle;
                    int hwndChild = 0;
                    Console.WriteLine("Excel application found - process id: {0}", excelId);

                    // Search the accessible child window (it has class name "EXCEL7") 
                    EnumChildCallback cb = new EnumChildCallback(EnumChildProc);
                    EnumChildWindows(hwnd, cb, ref hwndChild);

                    if (hwndChild != 0)
                    {
                        const uint OBJID_NATIVEOM = 0xFFFFFFF0;
                        Guid IID_IDispatch = new Guid("{00020400-0000-0000-C000-000000000046}");
                        ExcelWindow ptr;

                        int hr = AccessibleObjectFromWindow(hwndChild, OBJID_NATIVEOM, IID_IDispatch.ToByteArray(), out ptr);

                        if (hr >= 0)
                        {
                            // We successfully got a native OM IDispatch pointer, we can QI this for
                            // an Excel Application using reflection (and using UILanguageHelper to 
                            // fix http://support.microsoft.com/default.aspx?scid=kb;en-us;320369)
                            //
                            using (UILanguageHelper fix = new UILanguageHelper())
                            {
                                bool success = false;
                                while (success == false)
                                {
                                    try
                                    {
                                        xlApp = (Application)ptr.GetType().InvokeMember("Application", BindingFlags.GetProperty, null, ptr, null);
                                        success = true;
                                    }
                                    catch
                                    {
                                        SetForegroundWindow((IntPtr)hwnd);
                                        InputSimulator.SimulateKeyPress(VirtualKeyCode.EXECUTE);
                                    }
                                }
                                //object version = (Application)(xlApp.GetType().InvokeMember("Version", BindingFlags.GetField | BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, xlApp, null));
                                //Console.WriteLine(string.Format("Excel version is: {0}", version));
                            }


                            //xlApp = (Application)Marshal.GetActiveObject("Excel.Application"); // It selects only 1st excell app! 
                            xlBooks = xlApp.Workbooks;
                            var numBooks = xlBooks.Count;

                            Console.WriteLine("Number of opened workbooks: {0}", numBooks);

                            if (numBooks > 0)
                            {
                                xlBook = xlBooks[1];
                            }
                            if (numBooks == 0)
                            {
                                allExcelProcesses[i].Kill(); // kills process if there is no active workbooks (ghost-process)
                                Console.WriteLine("Empty process has been killed!");
                                break;
                            }

                            xlSheet = (Worksheet)xlBook.Worksheets[1];
                            string filePath = @"C:\ExcelSaver";

                            bool exists = System.IO.Directory.Exists(filePath);
                            if (!exists)
                                System.IO.Directory.CreateDirectory(filePath);

                            BringExcelWindowToFront(xlApp);
                            Console.WriteLine("Saving files: ");
                            foreach (Workbook wb in xlApp.Workbooks)
                            {
                                string fileName = wb.Name;
                                //string filePath = wb.Path;    default path (same as original file)
                                string fileNameAndPath = Path.Combine(filePath, fileName);
                                openedFiles.Add(fileNameAndPath);
                                Console.WriteLine("Saving {0}", fileName);

                                bool edit = true;
                                while (edit == true)
                                {
                                    edit = IsEditMode();
                                    if (edit == false)
                                    {
                                        xlApp.DisplayAlerts = false;
                                        wb.SaveAs(fileNameAndPath, wb.FileFormat, "", "", false, false,
                                          XlSaveAsAccessMode.xlExclusive, XlSaveConflictResolution.xlLocalSessionChanges);
                                        xlApp.DisplayAlerts = true;
                                    }
                                    else
                                    {
                                        InputSimulator.SimulateKeyPress(VirtualKeyCode.EXECUTE);
                                    }
                                }
                            }
                            xlBooks.Close();
                            xlApp.Quit();
                        }
                    }
                    else
                    {
                        allExcelProcesses[i].Kill();
                        Console.WriteLine("Empty process has been killed!");
                    }
                }
                Console.WriteLine("All files have been saved.");
                RemoveResources();
            }
            catch (Exception e)
            {
                Console.WriteLine("{0}", e.Message);
                Console.ReadKey();
            }
        }

        private bool IsEditMode()
        {
            object m = Type.Missing;
            const int MENU_ITEM_TYPE = 1;
            const int NEW_MENU = 18;

            // Get the "New" menu item.
            CommandBarControl oNewMenu = xlApp.CommandBars["Worksheet Menu Bar"].FindControl(MENU_ITEM_TYPE, NEW_MENU, m, m, true);

            if (oNewMenu != null)
            {
                // Check if "New" menu item is enabled or not.
                if (!oNewMenu.Enabled)
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// Open all saved excel Workbooks
        /// </summary>
        public void OpenAllFiles()
        {
            if (openedFiles.Count == 0)
            {
                return;
            }

            Console.WriteLine("\nOpening files:");
            xlApp = new Application();

            for (int i=0; i<openedFiles.Count; i++)
            {
                {
                    xlApp.Visible = false;
                    Console.WriteLine("Opening {0}", openedFiles[i]);
                    xlApp.Workbooks.Open(openedFiles[i], 0, false, 5, "", "", true, 
                        true, false, 0, false, 0, true, false, false);
                    openedFiles.Remove(openedFiles[i]);
                    i--;
                    xlApp.WindowState = XlWindowState.xlMinimized;
                    //   xlApp.WindowState = XlWindowState.xlMaximized; // excel runs in full-screen mode - without bars, controls... 
                }
            }

            Console.WriteLine("All files have been opened.");
            xlApp.Visible = true;


            RemoveResources();
        }

        /// <summary>
        /// Release memory
        /// </summary>
        public void RemoveResources()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();

            try
            {
                Marshal.FinalReleaseComObject(xlSheet);
            }
            catch (Exception ex)
            {
                //Console.WriteLine(ex);
                //Console.ReadKey();
            }

            try
            {
                foreach (Workbook wb in xlApp.Workbooks)
                {
                    Marshal.FinalReleaseComObject(wb);
                }
            }
            catch (Exception ex)
            {
                //Console.WriteLine(ex);
                //Console.ReadKey();
            }

            try
            {
                Marshal.FinalReleaseComObject(xlBooks);
            }
            catch (Exception ex)
            {
                //Console.WriteLine(ex);
                //Console.ReadKey();
            }

            try
            {
                Marshal.FinalReleaseComObject(xlApp);
            }
            catch (Exception ex)
            {
                //Console.WriteLine(ex);
                //Console.ReadKey();
            }
            Console.WriteLine("\nMemory has been released.");
        }

    }

    class UILanguageHelper : IDisposable
    {
        private CultureInfo _currentCulture;

        public UILanguageHelper()
        {
            // save current culture and set culture to en-US 
            _currentCulture = System.Threading.Thread.CurrentThread.CurrentCulture;
            System.Threading.Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
        }

        public void Dispose()
        {
            // reset to original culture 
            System.Threading.Thread.CurrentThread.CurrentCulture = _currentCulture;
        }
    }

}
