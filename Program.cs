using Microsoft.Office.Interop.Excel;
using System;
using System.Diagnostics;
using System.IO;
using System.Collections;
using System.Reflection;
using System.Collections.Generic;
using System.Configuration;

namespace RunMacro
{
    class Program
    {
        static Application ExcelApplication;
        static List<Workbook> wbList = new List<Workbook>();
        const string CheckFileName = "RunStart.ok";


        public static string AssemblyDirectory
        {
            get
            {
                string codeBase = Assembly.GetExecutingAssembly().CodeBase;
                UriBuilder uri = new UriBuilder(codeBase);
                string path = Uri.UnescapeDataString(uri.Path);
                return Path.GetDirectoryName(path);
            }
        }

        static void Main(string[] args)
        {
            bool visibility = false;
            try
            {
                if ((args == null) || (args.Length < 2))
                {
                    Console.WriteLine("Usage:");
                    Console.WriteLine("RunMacro.exe Open#\"<FileName1>\"");
                    Console.WriteLine("RunMacro.exe Open#\"<FileName2>\"");
                    Console.WriteLine("RunMacro.exe Open#\"<FileName3>\"");
                    Console.WriteLine("RunMacro.exe Run#\"<FileName4>\"#<MacroName>");
                    return;
                }

                if (File.Exists(Path.Combine(AssemblyDirectory, CheckFileName)))
                {
                    var msg = string.Format("Please delete the file {0} to proceed further!", Path.Combine(AssemblyDirectory, CheckFileName));
                    throw new Exception(msg);
                }
                else
                    File.Create(Path.Combine(AssemblyDirectory, CheckFileName));

                string commans = "";
                string fileName = "";
                string argument = "";

                try
                {
                    visibility = Convert.ToBoolean(ConfigurationManager.AppSettings["Visibility"]);
                }
                catch {/*Just make it as false*/ }


                ExcelApplication = new Application();
                ExcelApplication.Visible = visibility;
                ExcelApplication.DisplayAlerts = false;

                foreach (var arg in args)
                {
                    var values = arg.Split('#');
                    if (values.Length < 2)
                    {
                        Console.WriteLine("Usage:");
                        Console.WriteLine("RunMacro.exe Open#\"<FileName4>\"");
                        Console.WriteLine("Or");
                        Console.WriteLine("RunMacro.exe Run#\"<FileName4>\"#<MacroName>");
                        continue;
                    }

                    commans = arg.Split('#')[0];
                    fileName = arg.Split('#')[1];

                    if (values.Length > 2)
                    {
                        argument = arg.Split('#')[2];
                    }

                    switch (commans.Trim().ToUpper())
                    {
                        case "OPEN":
                            wbList.Add(OpenApplication(fileName));
                            break;
                        case "RUN":
                            RunMacroFromExcel(fileName, argument);
                            break;
                    }
                }
            }
            catch (Exception E)
            {
                Console.WriteLine(E.ToString());
                EventLog appLog = new EventLog();
                appLog.Source = "RunMacro";
                appLog.WriteEntry(E.ToString(), EventLogEntryType.Error);
            }
            finally
            {
                CloseApplication();
            }
        }

        static Workbook OpenApplication(string FilePath)
        {
            return GetWorkBook(FilePath);
        }

        /// <summary>
        /// There will be a run button in the given file. Run that.
        /// </summary>
        /// <param name="FilePath"></param>
        /// <returns></returns>
        static bool RunMacroFromExcel(string FilePath, string MacroName)
        {
            
            bool fileOpened = false;
            try
            {
                if (!File.Exists(FilePath))
                    throw new FileNotFoundException("File not found!", FilePath);

                wbList.Add(GetWorkBook(FilePath));
               
                ExcelApplication.Run(MacroName);

            }
            catch (Exception E)
            {
                Console.WriteLine(E.ToString());
            }
            return fileOpened;
        }

        static void CloseWorkBook(Workbook workbook)
        {
            try
            {
                workbook.Close();
            }
            catch { /* Do Nothing*/ }
        }

        /// <summary>
        /// Close excel application from memory
        /// </summary>
        static void CloseApplication()
        {
            if (ExcelApplication != null)
            {
                ExcelApplication.Visible = false;
                foreach (var wb in wbList)
                {
                    CloseWorkBook(wb);
                }

                if (ExcelApplication != null)
                {
                    ExcelApplication.Workbooks.Close();
                    ExcelApplication.Quit();
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(ExcelApplication);
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        static Workbook GetWorkBook(string FilePath)
        {
            if (!File.Exists(FilePath))
            {
                throw new Exception(string.Format("File {0} not found!", FilePath));
            }
            return ExcelApplication.Workbooks.Open(FilePath, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
        }

        static Worksheet GetWorkSheet(Workbook workbook, string sheetName)
        {
            bool found = false;
            if (workbook.Sheets.Count >= 0)
            {
                foreach (Worksheet worksheet in workbook.Worksheets)
                {
                    if (worksheet.Name == sheetName)
                    {
                        found = true;
                        break;
                    }
                }
            }

            if (found)
                return (Worksheet)workbook.Sheets[sheetName];
            else
                throw new Exception(string.Format("WorkSheet {0} not found!", sheetName));
        }
    }
}

