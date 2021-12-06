using System;
using System.Threading;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Diagnostics.Tracing;
using System.IO;
using System.Linq;
using System.Text;
using System.Reflection;
using CommandLine;

using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Outlook = Microsoft.Office.Interop.Outlook;
//using Microsoft.Office.Core;

using Workload;

namespace Helper
{
    class Utility
    {
        internal enum LogTimingFormat
        {
            LogGeoMeanTimingNormal,
            LogGeoMeanTimingWithMaxDrop,
            LogGeoMeanTimingWithMaxMinDrop,
            LogGeoMeanTimingWithMinDrop,
            LogGeoMeanMultiOperationTimingPerRepititon

        }

        internal enum TimingDropFormat
        {
            DropFirst,
            DropMax,
            DropMaxMin,
            DropMin

        }

        /// <summary>
        /// Minimizes console with the handler
        /// </summary>
        /// <param name="hWnd"></param>
        static void MinimizeConsole(IntPtr hWnd)
        {
            if (hWnd != IntPtr.Zero)
            {
                Logger.ShowWindow(hWnd, Logger.swMinimize);
            }
        }

        private static Word.Window FindDocumentWindow(Word.Application WordApp, Word._Document findThis)
        {
            foreach (Word.Window window in WordApp.Windows)
            {
                if (window.Document == findThis)
                {
                    return window;
                }
            }
            return null;

        }

        internal static void SetCaseID(string[] args, Logger.Settings set)
        {
            foreach (string argument in args)
            {
                if (argument.StartsWith("caseID"))
                {
                    set.caseID = Convert.ToInt32(argument.Split('-')[argument.Split('-').Length - 1]);
                }
            }
        }

        public static void ValidateResultsDirectory(Options opt)
        {
            string outputFoldername;
            if (Path.GetDirectoryName(opt.resultsDirectory) == "")
            {
                outputFoldername = Directory.GetCurrentDirectory() + "\\" + opt.resultsDirectory;
            }
            else
            {
                outputFoldername = Path.GetFullPath((opt.resultsDirectory));
            }
            opt.resultsDirectory = outputFoldername;
        }

        public static Dictionary<string, string> DefaultTopArguments()
        {
            Dictionary<string, string> defaultTopargs = new Dictionary<string, string>();
            defaultTopargs.Add("StartupPause", "2000");
            defaultTopargs.Add("Display", "1");
            defaultTopargs.Add("DisplayHeight", "700");
            defaultTopargs.Add("DisplayWidth", "1200");

            defaultTopargs.Add("runs", "1");
            defaultTopargs.Add("verbose", "True");
            defaultTopargs.Add("scriptversion", "none");
            defaultTopargs.Add("on-measure-start", "True");
            defaultTopargs.Add("on-measure-stop", "True");
            defaultTopargs.Add("results-directory", "..\\output");
            defaultTopargs.Add("TopHelp", "False");

            return defaultTopargs;
        }

        public static Dictionary<string, string> DefaultTopArgumentsPair()
        {
            Dictionary<string, string> defaultTopargsPair = new Dictionary<string, string>();

            defaultTopargsPair.Add("-r", "--runs");
            defaultTopargsPair.Add("-v", "--verbose");
            defaultTopargsPair.Add("-V", "--scriptversion");
            defaultTopargsPair.Add("-a","--on-measure-start");
            defaultTopargsPair.Add("-b", "--on-measure-stop");
            defaultTopargsPair.Add("-R","--results-directory");
            
            return defaultTopargsPair;

        }


        internal static string GetArgumentValues(string argName, string argument)
        {            
            
            string argumentValue = argument.Split('=')[argument.Split('=').Length - 1];
            Dictionary<string, string> defaultValues = DefaultTopArguments();
            if (argumentValue.Length == 0)
            {                
                return defaultValues[argName];
            }
            return argumentValue;
        }
        internal static void SetBatchScriptArguments(string[] args, ref Options options)
        {           

            foreach (string argument in args)
            {                
                if (argument.StartsWith("--StartupPause="))
                {
                    int value = Convert.ToInt32(GetArgumentValues("StartupPause", argument));
                    options.StartupPause = value;
                }
                if (argument.StartsWith("--Display="))
                {                    
                    options.Display = Convert.ToInt32(GetArgumentValues("Display", argument));                    
                }
                if (argument.StartsWith("--DisplayHeight="))
                {
                    options.DisplayHeight = Convert.ToDouble(GetArgumentValues("DisplayHeight", argument));
                }
                if (argument.StartsWith("--DisplayWidth="))
                {
                    options.DisplayWidth = Convert.ToDouble(GetArgumentValues("DisplayWidth", argument));
                }
                //
                if (argument.StartsWith("--runs="))
                {
                    options.runs = Convert.ToInt32(GetArgumentValues("runs", argument));
                }                
                if (argument.StartsWith("--scriptversion="))
                {
                    options.scriptversion = GetArgumentValues("scriptversion", argument);
                }                
                if (argument.StartsWith("--TopHelp="))
                {
                    bool checkHelp = false;
                    checkHelp = Convert.ToBoolean(GetArgumentValues("TopHelp", argument));
                    // does nothing
                }                
                if (argument.StartsWith("--verbose="))
                {
                    options.Verbose = GetArgumentValues("verbose", argument);
                }                
                if (argument.StartsWith("--on-measure-start="))
                {
                    options.onMeasureStart = GetArgumentValues("on-measure-start", argument);
                }                
                if (argument.StartsWith("--on-measure-stop="))
                {
                    options.onMeasureStop = GetArgumentValues("on-measure-stop", argument);
                }                
                if (argument.StartsWith("--results-directory="))
                {
                    options.resultsDirectory = GetArgumentValues("results-directory", argument);
                }                
            }
        }

        internal static bool CheckForDuplicates(List<string> inputArgs, string argument)
        {
            if (inputArgs.Contains(argument) && (argument.StartsWith("--") || argument.StartsWith("-")))
            {
                return false;
            }
            return true;
        }

        internal static string GetArgumentName(string argument)
        {
            if (argument.StartsWith("--") && argument.Length >= 3)
            {
                return argument.Split('=')[0];
            }

            if (argument.StartsWith("-") && argument.Length == 2)
            {
                return argument.Split('=')[0];
            }

            return argument;
        }
        internal static string RenameArgument(string argument)
        {
            Dictionary<string, string>  argsPair = DefaultTopArgumentsPair();
            if (argsPair.ContainsKey(argument))
            {
                return argsPair[argument];
            }
            else
            {
                return argument;
            }
            
        }
        internal static string[] CleanArguments(string[] args)
        {   
            List<string> cleandArgs = new List<string>();

            List<string> inputArgs = new List<string>();
            
            foreach (string argument in args)
            {
                // handles following cases
                // case 1: when top leavel arguments are not given 
                // case 2: when top level argument is not given and its given in custom case
                // case 3: when top level argument is given and its given in custom case

                // if argument starts and ends at -- and = remove it 
                if ((argument.StartsWith("--") && argument.EndsWith("=") )  )
                {
                    
                    continue;
                }
                else
                {
                    // handling single arguments 
                    // case 1: its is -r it will convert to --runs
                    string argumentMod = RenameArgument(argument);
                    // check if two time a argument is defined and ignore is its added in custom case : preference given to Top level script
                    
                    if (CheckForDuplicates(inputArgs, GetArgumentName(argumentMod)))
                    {
                        cleandArgs.Add(argumentMod);
                    }

                    inputArgs.Add(GetArgumentName(argumentMod));
                }

            }
                        
            return cleandArgs.ToArray();
        }
        public static void IterationEnd(string operationName, int iter, Logger.LogData logData, Logger.Settings set, int IterationPause)
        {
            // Add individual timings to the averageTime

            var elapsedS = set.iterationTimings[operationName].Last();

            Console.WriteLine("Run - " + set.repetition + ": Iteration - " + (iter + 1) + " - " + elapsedS);
            string[] operationNameWithRep = operationName.Split('_');
            string operationNameWithoutRep = operationNameWithRep[0];
            // Add logger timing
            logData.IterationTimings.Add(Logger.LogIterationTimings(iter: iter + 1, operationName: operationNameWithoutRep, value: elapsedS.ToString("0.000"), unit: "s", repetition: set.repetition));
            set.operationList.Add(operationNameWithoutRep);
            Thread.Sleep(IterationPause);
        }

        public static Logger.LogData LogTiming(Logger.Settings set, Options opt, Logger.LogData logData)
        {
            // Creating a dictionary with list as values
            // each key will contain timings of first iteration in each repetition
            List<string> operations = new List<string>(set.iterationTimings.Keys);
            double averageTimePerRepetition = 0.0;
            //double geoMeanTimePerRepetition = 1;
            int logRepeition = 0;
            foreach (var operationname in operations)
            {
                string[] operationNameWithRep = operationname.Split('_');
                string operationNameWithoutRep = operationNameWithRep[0];

                // CSV iteration logging
                for (int iter = 0; iter < set.iterationTimings[operationname].Count; iter++)
                {
                    if (!(set.iterationTimingsCollection.ContainsKey($"{operationNameWithoutRep}_iteration_{(iter + 1)}")))
                    {
                        set.iterationTimingsCollection.Add($"{operationNameWithoutRep}_iteration_{(iter + 1)}", new List<double>());
                    }
                }

                for (int iter = 0; iter < set.iterationTimings[operationname].Count; iter++)
                {
                    set.iterationTimingsCollection[$"{operationNameWithoutRep}_iteration_{(iter + 1)}"].Add(set.iterationTimings[operationname][iter]);
                    averageTimePerRepetition += set.iterationTimings[operationname][iter];

                }

                

                if (!(set.averageTime.ContainsKey(operationname)))
                {
                    set.averageTime.Add(operationname, averageTimePerRepetition / set.iterationTimings[operationname].Count);

                    string[] repetition = operationname.Split('_');

                    logRepeition = Convert.ToInt32(repetition[repetition.Length - 1]);

                    // Max iteration dropped in Jenka
                    Console.WriteLine("Runs - " + repetition[repetition.Length - 1] + ": Average_Time of " + operationNameWithoutRep + ":" + set.averageTime[operationname].ToString("0.000") + "s");

                    logData.AverageTime.Add(Logger.LogAverageTime(value: set.averageTime[operationname].ToString("0.000"),
                        operationName: operationNameWithoutRep, unit: "s", repetition: logRepeition));

                    // resetting values
                    averageTimePerRepetition = 0.0;

                }

            }

            // calculation of geo mean

            switch (set.timingFormat)
            {
                case LogTimingFormat.LogGeoMeanMultiOperationTimingPerRepititon:
                    logData = MultiOperationGeoMeanCalculation(set, opt, logData);
                    break;
                default:
                    // function call
                    logData = SingleOperationGeoMeanCalculation(set, opt, logData);
                    break;
            }

            

            return logData;
        }

        internal static Logger.LogData SingleOperationGeoMeanCalculation(Logger.Settings set, Options opt, Logger.LogData logData)
        {
            // SingleOperationGeoMean

            double geoMeanTimePerRepetition = 1;
            List<string> operations = new List<string>(set.iterationTimings.Keys);
            // Geo Mean calculation
            int logRepeition = 0;
            int totalIterationsPerRepittion = 0;
            int iterationOperationCounter = 0;

            foreach (var operationname in operations)
            {
                var targetList = GeoMeanCalculationFormat(set, set.iterationTimings[operationname],
                    out totalIterationsPerRepittion, opt);
                for (int iter = 0; iter < targetList.Count; iter++)
                {
                    geoMeanTimePerRepetition *= targetList[iter];
                }

                if ((iterationOperationCounter % totalIterationsPerRepittion) == 0)
                {
                    string[] repetition = operationname.Split('_');
                    logRepeition = Convert.ToInt32(repetition[repetition.Length - 1]);
                    string geoMeanOperation = $"GeoMean_rep_{logRepeition}";
                    geoMeanTimePerRepetition = Math.Pow(geoMeanTimePerRepetition, ((double)1 / (totalIterationsPerRepittion)));
                    set.GeoMean[geoMeanOperation] = geoMeanTimePerRepetition;
                    // resetting values

                    Console.WriteLine("Runs - " + logRepeition + ": GeoMean Time of " + geoMeanOperation + ":" + set.GeoMean[geoMeanOperation].ToString("0.000") + "s");

                    logData.GeoMean.Add(Logger.LogGeoMean(value: set.GeoMean[geoMeanOperation].ToString("0.000"),
                        operationName: geoMeanOperation, unit: "s", repetition: logRepeition));
                    geoMeanTimePerRepetition = 1;
                }
            }
            return logData;
        }

        internal static Logger.LogData MultiOperationGeoMeanCalculation(Logger.Settings set, Options opt, Logger.LogData logData)
        {

            Dictionary<int, Dictionary<string, List<Double>>> timingPerRepetitionperOperation = new Dictionary<int, Dictionary<string, List<Double>>>();
            List<string> iterationTimingsoperations = new List<string>(set.iterationTimings.Keys);

            // Make a list 
            foreach (var operationname in iterationTimingsoperations)
            {
                string[] repetitionName = operationname.Split('_');
                string currentOperationName = repetitionName[0];
                int repetitionNo = Convert.ToInt32(repetitionName[repetitionName.Length - 1]);
                if (!timingPerRepetitionperOperation.ContainsKey(repetitionNo))
                {
                    timingPerRepetitionperOperation.Add(repetitionNo, new Dictionary<string, List<Double>>());
                }
                if (!timingPerRepetitionperOperation[repetitionNo].ContainsKey(currentOperationName))
                {
                    timingPerRepetitionperOperation[repetitionNo].Add(currentOperationName, new List<Double>());
                }

                for (int iter = 0; iter < set.iterationTimings[operationname].Count; iter++)
                {
                    timingPerRepetitionperOperation[repetitionNo][currentOperationName].Add(set.iterationTimings[operationname][iter]);
                }

            }

            List<int> timingPerRepetitionperOperationOperations = new List<int>(timingPerRepetitionperOperation.Keys);

            double geoMeanTimePerRepetition = 1;

            foreach (int repNum in timingPerRepetitionperOperationOperations)
            {
                geoMeanTimePerRepetition = 1;
                int numberOfValues = 0;
                foreach (string operationname in timingPerRepetitionperOperation[repNum].Keys)
                {
                    // list of operation
                    var targetList = GeoMeanCalculationMultiOp(set, timingPerRepetitionperOperation[repNum][operationname], operationname, opt);

                    for (int iter = 0; iter < targetList.Count; iter++)
                    {
                        geoMeanTimePerRepetition *= targetList[iter];
                        numberOfValues += 1;
                    }

                }
                geoMeanTimePerRepetition = Math.Pow(geoMeanTimePerRepetition, ((double)1 / (numberOfValues)));

                Console.WriteLine("Runs - " + repNum + ": GeoMean Time :" + geoMeanTimePerRepetition.ToString("0.000") + "s");

                logData.GeoMean.Add(Logger.LogGeoMean(value: geoMeanTimePerRepetition.ToString("0.000"),
                    operationName: logData.Benchmark.name, unit: "s", repetition: repNum));
            }


            return logData;

        }
        internal static List<double> GeoMeanCalculationMultiOp(Logger.Settings set, List<double> inputList,
                string targetOperation, Options opt)
        {

            List<double> targetList = null;

            if (set.GeoMeanDropSetting.ContainsKey(targetOperation))
            {
                switch ((TimingDropFormat)set.GeoMeanDropSetting[targetOperation])
                {
                    case TimingDropFormat.DropFirst:
                        targetList = inputList;
                        targetList.RemoveAt(0);
                        break;
                    case TimingDropFormat.DropMax:
                        targetList = inputList.Where(x => x != inputList.Max()).ToList();
                        break;
                    case TimingDropFormat.DropMaxMin:
                        targetList = inputList.Where(x => x != inputList.Min() && x != inputList.Max()).ToList();
                        break;
                    case TimingDropFormat.DropMin:
                        targetList = inputList.Where(x => x != inputList.Min()).ToList();
                        break;

                }
            }
            else
            {
                targetList = inputList;
            }

            return targetList;
        }


        public static void CallStartBenchmark()
        {
            Logging.Log.StartBenchmark();
        }

        public static void CallEndBenchmark()
        {
            Logging.Log.EndBenchmark();
        }
        public static void ValidateOutputFiles(int rep, Options opt, Logger.LogData logData)
        {
            string outputFoldername;
            string outputFilename = Path.GetFileName(opt.OutputFileName);
            if (Path.GetDirectoryName(opt.OutputFileName) == "")
            {
                outputFoldername = Path.GetFullPath(opt.resultsDirectory);
            }
            else
            {
                outputFoldername = Path.GetFullPath(Path.GetDirectoryName(opt.OutputFileName));
            }

            // Check output directory exists
            if (Directory.Exists((outputFoldername)) == false)
            {
                Console.WriteLine("Output dumping folder doesn't exists. Program failed");
                Logger.ExceptionDeInit("Folder not found Exception", $"Output dumping folder {outputFoldername} does not exists", logData, opt);

            }
            opt.OutputFileName = Path.Combine(outputFoldername, outputFilename);

            //In some cases output is saved for each iteration
            for (int iter = 0; iter <= opt.Iterations; iter++)
            {
                if (File.Exists(GetFileName(opt.OutputFileName, rep, iter)) == true)
                {
                    File.Delete(GetFileName(opt.OutputFileName, rep, iter));
                    logData.Logging.Add(Logger.LogLogging(LogLevel: "Info", TimeStamp: DateTime.Now.ToString(), Detail: "Output file existed previously in same name is deleted"));
                }
            }
        }

        public static void ValidateInputFiles(Options opt, Logger.LogData logData)
        {
            string[] InputFileName = opt.InputFileName.Split(',');

            // Get input and output filenames
            for (int i = 0; i < InputFileName.Length; i++)
            {
                InputFileName[i] = Path.GetFullPath(InputFileName[i]);
                // Check if input file exists
                if (File.Exists(InputFileName[i]) == false)
                {
                    Console.WriteLine($"Input filename {Path.GetFileName(InputFileName[i])} doesn't exists. Program failed");
                    Logger.ExceptionDeInit("File not found Exception", $"Input filename {InputFileName[i]} does not exists", logData, opt);
                }
            }
            opt.InputFileName = InputFileName.Length > 1 ? string.Join(",", InputFileName) : InputFileName[0];
        }

        internal static List<double> GeoMeanCalculationFormat(Logger.Settings set, List<double> inputList,
                out int totalIterationsPerRepetition, Options opt)
        {
            List<double> targetList = null;
            totalIterationsPerRepetition = 0;
            List<string> operations = new List<string>(set.iterationTimings.Keys);
            foreach (var operationname in operations)
            {
                totalIterationsPerRepetition += set.iterationTimings[operationname].Count;
            }
            totalIterationsPerRepetition = totalIterationsPerRepetition / opt.runs;

            switch (set.timingFormat)
            {
                case LogTimingFormat.LogGeoMeanTimingNormal:
                    targetList = inputList;
                    totalIterationsPerRepetition -= 0;
                    break;
                case LogTimingFormat.LogGeoMeanTimingWithMaxDrop:
                    targetList = inputList.Where(x => x != inputList.Max()).ToList();
                    totalIterationsPerRepetition -= 1;
                    break;
                case LogTimingFormat.LogGeoMeanTimingWithMaxMinDrop:
                    targetList = inputList.Where(x => x != inputList.Min() && x != inputList.Max()).ToList();
                    totalIterationsPerRepetition -= 2;
                    break;
                case LogTimingFormat.LogGeoMeanTimingWithMinDrop:
                    targetList = inputList.Where(x => x != inputList.Min()).ToList();
                    totalIterationsPerRepetition -= 1;
                    break;

            }
            return targetList;
        }



        public static void StartFileOpen(string operationName, int rep, int iter, string statelabel)
        {

            string[] operationNameEvent = operationName.Split('_');
            operationName = operationNameEvent[0] + "_" + operationNameEvent[1];
            Logging.Log.Write($"{operationName}: {String.Format("{0:0000}", rep)}-{String.Format("{0:00}", iter)} {statelabel} ", new EventSourceOptions
            {
                Level = EventLevel.LogAlways,
                Opcode = EventOpcode.Start
            });
        }

        public static void StopFileOpen(string operationName, int rep, int iter, string statelabel)
        {

            string[] operationNameEvent = operationName.Split('_');
            operationName = operationNameEvent[0] + "_" + operationNameEvent[1];
            Logging.Log.Write($"{operationName}: {String.Format("{0:0000}", rep)}-{String.Format("{0:00}", iter)} {statelabel} ", new EventSourceOptions
            {
                Level = EventLevel.LogAlways,
                Opcode = EventOpcode.Stop
            });
        }

        internal static Tuple<double, double, double> Quartiles(double[] inputList)
        {
            int listLength = inputList.Length;
            int midValue = listLength / 2; // takes mid based on zero based index    
            double firstQuartile = 0;
            double secondQuartile = 0;
            double ThirdQuartile = 0;

            if (listLength % 2 == 0)
            {
                //making even for low and high point
                secondQuartile = (inputList[midValue - 1] + inputList[midValue]) / 2;
                int midValueMid = midValue / 2;

                //split
                if (midValue % 2 == 0)
                {
                    firstQuartile = (inputList[midValueMid - 1] + inputList[midValueMid]) / 2;
                    ThirdQuartile = (inputList[midValue + midValueMid - 1] + inputList[midValue + midValueMid]) / 2;
                }
                else
                {
                    firstQuartile = inputList[midValueMid];
                    ThirdQuartile = inputList[midValueMid + midValue];
                }
            }
            // case when listLength is 1
            else if (listLength == 1)
            {
                firstQuartile = inputList[0];
                secondQuartile = inputList[0];
                ThirdQuartile = inputList[0];
            }
            // case when listLength if odd
            else
            {
                secondQuartile = inputList[midValue];

                if ((listLength - 1) % 4 == 0)
                {
                    // (4n-1) points
                    int n = (listLength - 1) / 4;
                    firstQuartile = (inputList[n - 1] * .25) + (inputList[n] * .75);
                    ThirdQuartile = (inputList[3 * n] * .75) + (inputList[3 * n + 1] * .25);
                }
                else if ((listLength - 3) % 4 == 0)
                {
                    // (4n-3) points
                    int n = (listLength - 3) / 4;

                    firstQuartile = (inputList[n] * .75) + (inputList[n + 1] * .25);
                    ThirdQuartile = (inputList[3 * n + 1] * .25) + (inputList[3 * n + 2] * .75);
                }
            }

            return new Tuple<double, double, double>(firstQuartile, secondQuartile, ThirdQuartile);
        }

        // ValidateScreenDimensions
        public static void ValidateDisplayDimensions(Options opt, double sysheight, double syswidth)
        {
            if (opt.DisplayHeight > sysheight)
            {
                Console.WriteLine($"Given displayheight {opt.DisplayHeight} exceeds system display so setting to maxmimum system screen height");
                opt.DisplayHeight = sysheight;
            }
            if (opt.DisplayWidth > syswidth)
            {
                Console.WriteLine($"Given displaywidth {opt.DisplayHeight} exceeds system display so setting to maxmimum system screen width");
                opt.DisplayWidth = syswidth;
            }
        }

        //Application handler functions

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool SetForegroundWindow(IntPtr hWnd);

        public static void BringExcelWindowToFront(Excel.Application xlApp)
        {
            SetForegroundWindow((IntPtr)xlApp.Hwnd);
        }

        public static void BringWordWindowToFront(Word.Window window)
        {
            SetForegroundWindow((IntPtr)window.Hwnd);
        }

        public static void BringPowerpointWindowToFront(PowerPoint.Application pptApp)
        {
            SetForegroundWindow((IntPtr)pptApp.HWND);
        }

        /// Opens Excel app , minimises console and then opens workbook .Triggers ETW events for opening file

        public static void ExcelInit(out Excel.Application app, out List<Excel._Workbook> workbooks, out List<Excel._Worksheet> worksheets,
            string operationName, Logger.Settings set, Options opt)
        {
            // Open the Excel Application
            app = new Excel.Application
            {
                Visible = true,
                DisplayAlerts = false
            };

            set.previousFullScreenSetting = Convert.ToInt32(app.DisplayFullScreen);

            if (opt.Display == 1)
            {
                app.WindowState = Excel.XlWindowState.xlMinimized;
                app.WindowState = Excel.XlWindowState.xlMaximized;

                // Turning to full screen mode
                app.DisplayFullScreen = true;

            }
            else if (opt.Display == 2)
            {


                app.WindowState = Excel.XlWindowState.xlNormal;
                set.originalScreenSetting.height = (double)app.Height;
                set.originalScreenSetting.width = (double)app.Width;

                Logger.ScreenResolution screenDimenstion = Logger.GetScreenDimension();
                ValidateDisplayDimensions(opt, screenDimenstion.height, screenDimenstion.width);

                app.Height = opt.DisplayHeight;
                app.Width = opt.DisplayWidth;

            }
            BringExcelWindowToFront(app);


            //MinimizeConsole(set.hWnd);
            Thread.Sleep(opt.SeparationPause);
            workbooks = new List<Excel._Workbook>();
            worksheets = new List<Excel._Worksheet>();

            if (set.openFileFromInit)
            {
                int numberOfWorkBooks = set.fileNames.Count;

                for (int workbookNumber = 0; workbookNumber < numberOfWorkBooks; workbookNumber++)
                {
                    // ETW logging for opening a file
                    StartFileOpen(operationName, set.repetition, opt.Iterations,
                        $"Excel Start File Open {(workbookNumber + 1)}_{set.repetition}");
                    // Open Input Excel workbook
                    workbooks.Add(app.Workbooks.Open(set.fileNames[workbookNumber], UpdateLinks: 0, ReadOnly: false));

                    // ETW logging for opening a file            
                    StopFileOpen(operationName, set.repetition, opt.Iterations,
                        $"Excel Stop File Open {(workbookNumber + 1)}_{set.repetition}");

                    Thread.Sleep(opt.SeparationPause);
                }

                // Activate workbook
                workbooks[0].Activate();

            }

            Thread.Sleep(opt.StartupPause);
        }


        /// Closes the file step by step. First it restores the original screensettings and then closes the workbook and app

        public static void ExcelDeInit(Excel.Application app, List<Excel._Workbook> workbooks, List<Excel._Worksheet> worksheets, Logger.Settings set, Options opt)
        {
            if (set.closeFileFromDeInit)
            {
                // Restoring original setting
                GC.Collect();
                GC.WaitForPendingFinalizers();
                // Restoring original setting
                if (app != null)
                {
                    if ((set.originalScreenSetting.height != 0) && (set.originalScreenSetting.width != 0))
                    {
                        app.WindowState = Excel.XlWindowState.xlNormal;
                        app.Height = set.originalScreenSetting.height;
                        app.Width = set.originalScreenSetting.width;
                    }

                    app.DisplayFullScreen = Convert.ToBoolean(set.previousFullScreenSetting);

                }

                if (worksheets != null)
                {
                    for (int sheet = 0; sheet < worksheets.Count; sheet++)
                    {
                        if (worksheets[sheet] != null)
                        {
                            Marshal.ReleaseComObject(worksheets[sheet]);
                        }
                    }
                }

                if (workbooks != null)
                {
                    for (int workbook = 0; workbook < workbooks.Count; workbook++)
                    {

                        // Close workbook
                        if (workbooks[workbook] != null)
                        {
                            workbooks[workbook].Close(false, Type.Missing, Type.Missing);
                            Thread.Sleep(opt.SeparationPause);
                            Marshal.ReleaseComObject(workbooks[workbook]);
                        }
                    }
                }

                // Close Excel application
                if (app != null)
                {
                    app.Quit();
                    Marshal.ReleaseComObject(app);
                }
            }
        }

        public static void PowerPointInit(out PowerPoint.Application app, out List<PowerPoint.Presentation> presentations, out List<PowerPoint.Slide> slides, string operationName, Logger.Settings set, Options opt, Logger.LogData logData)
        {
            // Open PowerPoint Application
            app = new PowerPoint.Application
            {
                DisplayAlerts = PowerPoint.PpAlertLevel.ppAlertsNone,
                //Visible = MsoTriState.msoTrue
            };

            set.previousFullScreenSetting = (int)app.WindowState;

            if (opt.Display == 1)
            {
                // Making PowerPoint to run in Foreground
                app.WindowState = PowerPoint.PpWindowState.ppWindowMinimized;
                app.WindowState = PowerPoint.PpWindowState.ppWindowMaximized;
            }
            else if (opt.Display == 2)
            {

                app.WindowState = PowerPoint.PpWindowState.ppWindowNormal;
                // storing previous setting
                set.originalScreenSetting.height = (float)app.Height;
                set.originalScreenSetting.width = (float)app.Width;

                Logger.ScreenResolution screenDimenstion = Logger.GetScreenDimension();
                ValidateDisplayDimensions(opt, screenDimenstion.height, screenDimenstion.width);

                app.Height = (float)opt.DisplayHeight;
                app.Width = (float)opt.DisplayWidth;
            }
            BringPowerpointWindowToFront(app);

            //MinimizeConsole(set.hWnd);

            Thread.Sleep(opt.SeparationPause);

            presentations = new List<PowerPoint.Presentation>();
            slides = new List<PowerPoint.Slide>();

            if (set.openFileFromInit)
            {
                int numberOfPresentations = set.fileNames.Count;


                for (int presentationNumber = 0; presentationNumber < numberOfPresentations; presentationNumber++)
                {
                    StartFileOpen(operationName, set.repetition, opt.Iterations, $"Powerpoint Start File Open {(presentationNumber + 1)}_{set.repetition}");
                    // Open Input Excel workbook
                    presentations.Add(app.Presentations.Open(set.fileNames[presentationNumber]));
                    StopFileOpen(operationName, set.repetition, opt.Iterations, $"Powerpoint Stop File Open {(presentationNumber + 1)}_{set.repetition}");
                    Thread.Sleep(opt.SeparationPause);
                }
            }
            Thread.Sleep(opt.StartupPause);


        }


        public static void PowerPointDeInit(PowerPoint.Application app, List<PowerPoint.Presentation> presentations, List<PowerPoint.Slide> slides, Logger.Settings set, int ThreadSleep)
        {
            if (set.closeFileFromDeInit)
            {
                if (app != null)
                {
                    if ((set.originalScreenSetting.height != 0) && (set.originalScreenSetting.width != 0))
                    {
                        app.WindowState = PowerPoint.PpWindowState.ppWindowNormal;
                        app.Height = (float)set.originalScreenSetting.height;
                        app.Width = (float)set.originalScreenSetting.width;
                    }

                    app.WindowState = (PowerPoint.PpWindowState)set.previousFullScreenSetting;

                }

                if (slides != null)
                {
                    for (int slide = 0; slide < slides.Count; slide++)
                    {
                        if (slides[slide] != null)
                        {
                            Marshal.ReleaseComObject(slides[slide]);
                        }
                    }
                }

                if (presentations != null)
                {
                    for (int presentation = 0; presentation < presentations.Count; presentation++)
                    {

                        // Close presentation
                        if (presentations[presentation] != null)
                        {
                            presentations[presentation].Close();
                            Thread.Sleep(ThreadSleep);
                            Marshal.ReleaseComObject(presentations[presentation]);
                        }
                    }
                }

                // Close Power Point application
                if (app != null)
                {
                    app.Quit();
                    Marshal.ReleaseComObject(app);
                }
                Thread.Sleep(ThreadSleep);
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

        }
        public static void WordInit(out Word.Application app, out List<Word._Document> documents, out List<Word.Window> docWindows,
                string operationName, Logger.Settings set, Options opt, Logger.LogData logData)
        {
            // Open the Word Application
            app = new Word.Application
            {
                Visible = true,
            };

            object readOnly = false;
            object isVisible = true;
            object missing = System.Reflection.Missing.Value;

            set.previousFullScreenSetting = Convert.ToInt32(app.WindowState);



            //MinimizeConsole(set.hWnd);
            if (opt.Display == 1)
            {
                // Turning to full screen mode
                app.WindowState = Word.WdWindowState.wdWindowStateMinimize;
                app.WindowState = Word.WdWindowState.wdWindowStateMaximize;
            }
            else if (opt.Display == 2)
            {
                app.WindowState = Word.WdWindowState.wdWindowStateNormal;
                set.originalScreenSetting.height = (double)app.Height;
                set.originalScreenSetting.width = (double)app.Width;

                Logger.ScreenResolution screenDimenstion = Logger.GetScreenDimension();
                ValidateDisplayDimensions(opt, screenDimenstion.height, screenDimenstion.width);
                app.Height = (int)opt.DisplayHeight;
                app.Width = (int)opt.DisplayWidth;
            }
            
            Thread.Sleep(opt.SeparationPause);


            docWindows = new List<Word.Window>();
            documents = new List<Word._Document>();

            if (set.openFileFromInit)
            {

                int numberOfDocuments = set.fileNames.Count;

                for (int documentNumber = 0; documentNumber < numberOfDocuments; documentNumber++)
                {
                    // ETW logging for opening a file
                    StartFileOpen(operationName, set.repetition, opt.Iterations, $"Word Start File Open {(documentNumber + 1)}_{set.repetition}");
                    // Open Input Excel workbook
                    documents.Add(app.Documents.Open(set.fileNames[documentNumber], ref missing,
                                  ref readOnly, ref missing, ref missing,
                                  ref missing, ref missing, ref missing,
                                  ref missing, ref missing, ref missing,
                                  ref isVisible, ref missing, ref missing,
                                  ref missing, ref missing));

                    // ETW logging for opening a file            
                    StopFileOpen(operationName, set.repetition, opt.Iterations, $"Word Stop File Open {(documentNumber + 1)}_{set.repetition}");

                    Thread.Sleep(opt.SeparationPause);
                }
                
                Thread.Sleep(opt.SeparationPause);

                // Activate document
                documents[0].Activate();


                for (int documentNumber = 0; documentNumber < numberOfDocuments; documentNumber++)
                {
                    docWindows.Add(FindDocumentWindow(app, documents[documentNumber]));

                    if (docWindows[documentNumber] == null)
                    {
                        Console.WriteLine("Cannot get a reference to the destination document window!");
                        Logger.ExceptionDeInit("Exception: Destination document window not found", "Cannot get a reference to the destination document window!", logData, opt);
                    }
                    docWindows[documentNumber].Panes[1].View.Zoom.Percentage = 100;
                    //set.previousFullScreenSetting = Convert.ToInt32(docWindows[0].View.FullScreen);
                    if (opt.Display == 1)
                    {
                        docWindows[documentNumber].Panes[1].View.Zoom.Percentage = 100;
                        docWindows[documentNumber].View.FullScreen = true;
                    }
                    BringWordWindowToFront(docWindows[documentNumber]);
                }
                Thread.Sleep(opt.SeparationPause);
            }
            Thread.Sleep(opt.StartupPause);
        }


        public static void WordDeInit(Word.Application app, List<Word._Document> documents, Logger.Settings set, int ThreadSleep)
        {
            if (set.closeFileFromDeInit)
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
                // Restoring original setting
                if (app != null)
                {
                    if ((set.originalScreenSetting.height != 0) && (set.originalScreenSetting.width != 0))
                    {
                        app.WindowState = Word.WdWindowState.wdWindowStateNormal;
                        app.Height = (int)set.originalScreenSetting.height;
                        app.Width = (int)set.originalScreenSetting.width;
                    }

                    app.WindowState = (Word.WdWindowState)Convert.ToInt32(set.previousFullScreenSetting);

                }


                if (documents != null)
                {
                    for (int doc = 0; doc < documents.Count; doc++)
                    {
                        if (documents[doc] != null)
                        {
                            documents[doc].Close(false, Type.Missing, Type.Missing);
                            Thread.Sleep(ThreadSleep);
                            Marshal.ReleaseComObject(documents[doc]);
                        }
                    }
                }

                // Close Word application
                if (app != null)
                {
                    app.Quit();
                    Marshal.ReleaseComObject(app);
                }
            }
        }

		public static void OutlookInit(out Outlook.Application app, out Outlook.NameSpace nameSpace, Outlook.OlDefaultFolders defaultFolder, Logger.Settings set, Options opt)
        {

            app = new Outlook.Application();
            nameSpace = app.GetNamespace("MAPI");
            nameSpace.Logon("Outlook", Missing.Value, false, true);

            // Thread Sleep for opening outlook app
            Thread.Sleep(opt.SeparationPause);

            Outlook.MAPIFolder folder = nameSpace.GetDefaultFolder(defaultFolder);
            folder.Display();

            //variable lines
            Thread.Sleep(opt.SeparationPause);

            set.previousFullScreenSetting = Convert.ToInt32(app.ActiveWindow().WindowState);

            if (opt.Display == 1)
            {
                // Turing to full screen mode
                app.ActiveWindow().WindowState = Outlook.OlWindowState.olMinimized;
                app.ActiveWindow().WindowState = Outlook.OlWindowState.olMaximized;
            }
            else if (opt.Display == 2)
            {
                app.ActiveWindow().WindowState = Outlook.OlWindowState.olNormalWindow;
                set.originalScreenSetting.height = (double)app.ActiveWindow().Height;
                set.originalScreenSetting.width = (double)app.ActiveWindow().Width;

                Logger.ScreenResolution screenDimenstion = Logger.GetScreenDimension();
                ValidateDisplayDimensions(opt, screenDimenstion.height, screenDimenstion.width);


                app.ActiveWindow().Height = opt.DisplayHeight;
                app.ActiveWindow().Width = opt.DisplayWidth;

                MinimizeConsole(set.hWnd);
            }
            app.ActiveWindow().Activate();

            //MinimizeConsole(set.hWnd);
            Thread.Sleep(opt.StartupPause);
        }

        public static void OutlookDeInit(Outlook.Application app, Outlook.NameSpace nameSpace, Outlook.MAPIFolder importFolder, Logger.Settings set, int ThreadSleep)
        {
            if (app != null)
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
                if ((set.originalScreenSetting.height != 0) && (set.originalScreenSetting.width != 0))
                {
                    app.ActiveWindow().WindowState = Outlook.OlWindowState.olNormalWindow;
                    app.ActiveWindow().Height = set.originalScreenSetting.height;
                    app.ActiveWindow().Width = set.originalScreenSetting.width;

                    Logger.ShowWindow(set.hWnd, Logger.swShowNormal);
                }

                app.ActiveWindow().WindowState = (Outlook.OlWindowState)set.previousFullScreenSetting;

            }
			
            if (importFolder != null)
                app.Session.RemoveStore(importFolder);

            // Close Outlook application
            if (app != null)
            {
                app.Quit();
                Thread.Sleep(ThreadSleep);
                Marshal.ReleaseComObject(app);
                Marshal.ReleaseComObject(nameSpace);
            }
			//Logger.ShowWindow(set.hWnd, Logger.swShowNormal);
        }

        public static void RemoveDataCopy(int rep, Options opt, Logger.LogData logData)
        {
            string outputFoldername;
            string outputFilename = Path.GetFileName(opt.InputFileName);
            if (Path.GetDirectoryName(opt.InputFileName) == "")
            {
                outputFoldername = Directory.GetCurrentDirectory();
            }
            else
            {
                outputFoldername = Path.GetFullPath(Path.GetDirectoryName(opt.InputFileName));
            }

            // Check output directory exists
            if (Directory.Exists((outputFoldername)) == false)
            {
                Console.WriteLine("Output dumping folder doesn't exists. Program failed");
                Logger.ExceptionDeInit("Folder not found Exception", $"Output dumping folder {outputFoldername} does not exists", logData, opt);

            }
            opt.InputFileName = Path.Combine(outputFoldername, outputFilename);

            if (File.Exists(Path.Combine(Path.GetDirectoryName(opt.InputFileName), $"Datafile_{rep}.pst")) == true)
            {
                File.Delete(Path.Combine(Path.GetDirectoryName(opt.InputFileName), $"Datafile_{rep}.pst"));
                logData.Logging.Add(Logger.LogLogging(LogLevel: "Info", TimeStamp: DateTime.Now.ToString(), Detail: "Output file existed previously in same name is deleted"));
            }
        }


        /// Gets the file name given by user and concatenates repetation number and iteration number with it
        /// Example: BenchmakrResult099910.xlsx, where 0999 -> zero padded repetition number , 10 -> iteration number        
        public static string GetFileName(string outputFileName , int rep, int iter)
        {
            string filename = Path.GetFileNameWithoutExtension(outputFileName);
            string extension = Path.GetExtension(outputFileName);
            string reps = String.Format("{0:0000}", rep);
            string iteration = String.Format("{0:00}", iter);
            string rep_outputFileName = Path.GetDirectoryName(outputFileName) + "\\"+ filename + reps + iteration + extension; 

            return rep_outputFileName;
        }

        //Gets the extension from the given inputfile some programs uses multiple input files
        public static string GetFileName(string outputFileName, int rep, int iter, string inputFileName)
        {
            string rep_outputFileName = Path.GetDirectoryName(outputFileName) + "\\" + Path.GetFileNameWithoutExtension(outputFileName) + Path.GetExtension(inputFileName);
            return GetFileName(rep_outputFileName, rep, iter);
        }
    }

}
