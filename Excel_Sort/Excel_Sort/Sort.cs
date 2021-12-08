using System;
using System.Collections.Generic;
using System.Threading;
using System.Diagnostics;
using System.IO;

using Excel = Microsoft.Office.Interop.Excel;

using Helper;
using CommandLine;

namespace Workload
{
    public class Options : CommonOptions
    {
        [Option("SheetNumber", Required = false, HelpText = "Sheet number eg: 1", Default = 1)]
        public int SheetNumber { get; set; }

        [Option("SortOrder", Required = false, HelpText = "Sort order : ASC or DES", Default = "ASC")]
        public string SortOrder { get; set; }

        [Option("Range", Required = false, HelpText = "Range in which Sort to be perfromed eg: A1:V70000 ", Default = "A1:V70000")]
        public string Range { get; set; }

        [Option("ColumnList", Required = false, HelpText = "column numbers in the format of col1,col2,col3,col4. Example: --ColumnList 1,2,3", Default = "1,5,18,3,15,7,4,12,9,16")]
        public string ColumnList { get; set; }

        public const bool inputFileFlag = true;
        public const bool outputFileFlag = true;
    }

    class SpecificArguments
    {
        // Command line option for Sort
        public static Options ParseArgument(string[] args, Logger.Settings set)
        {
            Options options = new Options();

            if (args.Length != 0)
            {
                Utility.SetCaseID(args, set);

                // Invoke Sort default
                if (args[0] == "default")
                {
                    options.InputFileName = "..\\input\\MOCK_Data_Only_for_sorting.xlsx";
                    options.OutputFileName = "SortResult.xlsx";
                    options.Iterations = 10;
                    options.SheetNumber = 1;
                    options.SortOrder = "ASC";
                    options.Range = "A1:V70000";
                    options.ColumnList = "1,5,18,2,15,7,4,12,9,16";
                }
            }
            Arguments.ParseArgument(args, ref options);
            return options;
        }
    }

    static class Sort
    {
        static void CreateCSV(Options opt, Logger.LogData logData) 
        {
            string status = "success";
            string operationname = logData.Benchmark.name;
            string startTime = logData.StartTime.time;
            string sysPath = Path.GetFullPath(opt.resultsDirectory);
            string csvfile = sysPath + "\\" + operationname + "_" + startTime + ".csv";

            string[] lines = File.ReadAllLines(csvfile);

            if (lines.Length == 0)
            {
                status = "failure";
            }
            string[] columns = opt.ColumnList.Split(',');
            //add new column to the header row
            lines[0] += ",InputFileName,SheetNumber,ColumnID,SortOrder,Range";

            //add new column value for each row.
            for (int i = 1; i < lines.Length; i++)
            {
                lines[i] += "," + Path.GetFileName(opt.InputFileName) + "," + opt.SheetNumber + "," + columns[i-1] + "," + opt.SortOrder + "," + opt.Range;
            }

            //write the new content
            File.WriteAllLines(csvfile, lines);
            Logger.LogProgramEnd(status,opt);
        }

        // Sort Validations
        static void ValidateColumnList(out string[] columnList, Logger.LogData logData, Options opt)
        {
            columnList = opt.ColumnList.Split(',');

            // Check the total number of iterations and columns number 
            if (opt.Iterations == -1)
            {
                opt.Iterations = columnList.Length;
            }           
            else if (opt.Iterations > columnList.Length)
            {
                string[] modColumnList = new string[opt.Iterations];
                for (int i = 0; i < opt.Iterations; i++)
                {                    
                    modColumnList[i] = columnList[(i%columnList.Length)];
                }
                columnList = modColumnList;
                opt.ColumnList = string.Join(",", columnList);
            }
            else if (opt.Iterations < columnList.Length)
            {
                string[] modColumnList = new string[opt.Iterations];
                for (int i = 0; i < opt.Iterations; i++)
                {                    
                    modColumnList[i] = columnList[i];
                }
                columnList = modColumnList;
                opt.ColumnList = string.Join(",", columnList);
            }

        }

        static void ValidateSheetNumber(List<Excel._Workbook> workbooks, List<Excel._Worksheet> worksheets, Logger.LogData logData, int SheetNumber)
        {
            int numSheets = workbooks[0].Sheets.Count;
            if (SheetNumber == 0 || SheetNumber > numSheets)
            {
                Console.WriteLine("Please enter valid sheet number");
                logData.Logging.Add(Logger.LogLogging(LogLevel: "Error", TimeStamp: DateTime.Now.ToString(), Detail: "Please enter valid sheet number"));
                throw new IndexOutOfRangeException("Please enter valid sheet number");
            }

        }

        // Function Description: Sort the Columns in Excel Worksheet 
        static void ExcelSort(Excel.Application app, List<Excel._Workbook> workbooks, List<Excel._Worksheet> worksheets, Logger.Settings set,
                                Options opt, Logger.LogData logData, int repeat)
        {

            for (int rep = 1; rep <= repeat; rep++)
            {
                string operationName = "Sort_" + rep.ToString();
                set.repetition = rep;   
                string[] columnList = null;
                ValidateColumnList(out columnList, logData, opt);
                Utility.ValidateOutputFiles(rep, opt, logData);
                set.fileNames.Add(opt.InputFileName);

                try
                {
                    Utility.ExcelInit(out app, out workbooks, out worksheets, operationName, set, opt);
                    ValidateSheetNumber(workbooks, worksheets, logData, opt.SheetNumber);

                    // Adding worksheet to worksheet list and activating
                    worksheets.Add(workbooks[0].Sheets[opt.SheetNumber]);
                    worksheets[0].Activate();

                    Excel.Range range = worksheets[0].Range[opt.Range];
                    range.Select();
                    set.iterationTimings.Add(operationName, new List<double>());

                    Thread.Sleep(opt.SeparationPause);

                    for (int iter = 0; iter < opt.Iterations; iter++)
                    {
                        // Get Column number.
                        int colsId = Int32.Parse(columnList[iter]);
                        Console.WriteLine($"Sorting for Column ID {colsId}");
                        Excel.XlSortOrder sortOrder = 0;
                        Stopwatch watch = new Stopwatch();

                        if (opt.SortOrder == "ASC")
                        {
                            sortOrder = Excel.XlSortOrder.xlAscending;
                        }
                        else if (opt.SortOrder == "DES")
                        {
                            sortOrder = Excel.XlSortOrder.xlDescending;
                        }

                        Logger.IterationEventStart(operationName, rep, iter);

                        // Sort the Selected range.  
                        watch.Start();
                        range.Sort(range.Columns[colsId], sortOrder, Header: Excel.XlYesNoGuess.xlYes);
                        watch.Stop();

                        Logger.IterationEventEnd(operationName, rep, iter);

                        set.iterationTimings[operationName].Add(watch.ElapsedMilliseconds / 1000.0);
                        Utility.IterationEnd(operationName, iter, logData, set, opt.IterationPause);
                    }

                    Thread.Sleep(opt.SeparationPause);

                    // Save the workbook.
                    
                    workbooks[0].SaveAs(Utility.GetFileName(opt.OutputFileName, rep, opt.Iterations));
                    Thread.Sleep(opt.SeparationPause);

                    Utility.ExcelDeInit(app, workbooks, worksheets, set, opt);

                    set.status = "success";
                    Thread.Sleep(opt.SeparationPause);
                    

                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception raised. Program failed");
                    Console.WriteLine(e);
                    Console.WriteLine(e.Message);

                    Utility.ExcelDeInit(app, workbooks, worksheets, set, opt);
                    set.exception = e;
                    set.status = "failure";
                    Thread.Sleep(opt.SeparationPause);
                    return;

                }
                app = null;
                workbooks = null;
                worksheets = null;
                set.fileNames = new List<string>();
            }
            set.timingFormat = Utility.LogTimingFormat.LogGeoMeanTimingNormal;
            Logger.DeInit(logData, set, set.status, opt, set.exception);
            CreateCSV( opt, logData);
        }

        /// The main entry point for the application.
        [STAThread]
        static void Main(string[] args)
        {
            Options opt;

            Logger.Settings set = new Logger.Settings();
            opt = SpecificArguments.ParseArgument(args, set);

            if (args.Length == 0 || opt.InputFileName == null)
            {
                return;
            }

            string benchMarkTestName = "Excel_Sort";
            Logger.LogData logData = Logger.Init(opt, benchMarkTestName, "Excel");
            Console.WriteLine("Sorting in " + opt.SortOrder);
            Utility.ValidateInputFiles(opt, logData);

            // Decalre the Excel objects
            Excel.Application app = null;
            List<Excel._Workbook> workbooks = null;
            List<Excel._Worksheet> worksheets = null;

            ExcelSort(app, workbooks, worksheets, set,  opt, logData, (opt.runs));
        }

    }
}