using System;
using System.Collections.Generic;
using System.Web.Script.Serialization;
using System.IO;
using System.Text;
using System.Diagnostics;
using System.Diagnostics.Tracing;
using Microsoft.Win32;
using System.Threading;
using System.Runtime.InteropServices;
using System.Management;
using Workload;

namespace Helper
{
    class Logger
    {
        public const int swMinimize = 6; // Minimize console
        public const int swShowNormal = 1; // Restore normal console

        [DllImport("Kernel32.dll")]
        public static extern IntPtr GetConsoleWindow();

        [DllImport("User32.dll")]
        public static extern bool ShowWindow(IntPtr hWnd, int cmdShow);
        public class BenchmarkStruct
        {
            public string name { get; set; }
            public string version { get; set; }
            public string msOfficeVersion { get; set; }
        }
        public class LoggingStruct
        {
            public string logLevel { get; set; }
            public string timeStamp { get; set; }
            public string detail { get; set; }
        }

        public class IterationTimingsStruct
        {
            public int iteration { get; set; }
            public string operationName { get; set; }
            public string value { get; set; }
            public string unit { get; set; }
            public int repetition { get; set; }

        }
        public class AverageTimeStruct
        {
            public string value { get; set; }
            public string operationName { get; set; }
            public string unit { get; set; }
            public int repetition { get; set; }
        }

        public class GeoMeanStruct
        {
            public string value { get; set; }
            public string operationName { get; set; }
            public string unit { get; set; }
            public int repetition { get; set; }
        }

        public class StartTimeStruct
        {
            public string msSinceEpoch { get; set; }
            public string time { get; set; }
        }
        public class StopTimeStruct
        {
            public string msSinceEpoch { get; set; }
            public string time { get; set; }
        }
        public class EtwInfoStruct
        {
            public string[] events { get; set; }
            public int count { get; set; }
        }
        public class ExceptionStruct
        {
            public string name { get; set; }
            public string message { get; set; }
        }

        public class StatusStruct
        {
            public string statusCode { get; set; }
            public string detailedMessage { get; set; }
        }

        public class BoardInfoStruct
        {
            public string manufaturer { get; set; }
            public string serialNumber { get; set; }
            public string productID { get; set; }
            public string version { get; set; }
        }
        public class GraphicsDriverStruct
        {
            public string videoProcessor { get; set; }
            public string videoModeDescription { get; set; }

        }

        public class MemoryConfigStruct
        {
            public string totalMemoryAvailable { get; set; }

        }
        public class DiskConfigStruct
        {
            public string availableFreeSpaceCurrentUser { get; set; }
            public string aotalFreeSpace { get; set; }
            public string aotalSize { get; set; }

        }

        public class SystemInfoStruct
        {
            public string systemName { get; set; }
            public string processorModel { get; set; }
            public Dictionary<string, string> boardInfo { get; set; }
            public Dictionary<string, string> graphicsDriver { get; set; }
            public Dictionary<string, string> memoryConfig { get; set; }
            public Dictionary<string, string> diskConfig { get; set; }
            public string osVersion { get; set; }
            public Dictionary<string, string> biosVersion { get; set; }
            public Dictionary<string, string> batteryInfo { get; set; }

        }

        public class ConditionsStruct
        {
            public string detailedMessage { get; set; }
        }

        public class BehaviourStruct
        {
            public string detailedMessage { get; set; }
        }
        public class ScreenResolution
        {
            public double height { get; set; }
            public double width { get; set; }
        }

        public class LogData
        {
            public BenchmarkStruct Benchmark { get; set; }
            public List<LoggingStruct> Logging { get; set; }
            public object Parameters { get; set; }
            public List<IterationTimingsStruct> IterationTimings { get; set; }
            public List<AverageTimeStruct> AverageTime { get; set; }
            public List<GeoMeanStruct> GeoMean { get; set; }
            public StartTimeStruct StartTime { get; set; }
            public StopTimeStruct StopTime { get; set; }
            public EtwInfoStruct EtwInfo { get; set; }
            public ExceptionStruct Exception { get; set; }
            public SystemInfoStruct SystemInfo { get; set; }
            public ConditionsStruct Conditions { get; set; }
            public BehaviourStruct Behaviour { get; set; }
            public StatusStruct Status { get; set; }

        }

        public static BenchmarkStruct LogBenchMark(string benchMarkName, string version, string msofficeVersion)
        {
            BenchmarkStruct benchmarkStructVar = new BenchmarkStruct { name = benchMarkName, version = version, msOfficeVersion = msofficeVersion };
            return benchmarkStructVar;
        }

        public static string OfficeInformation(string workloadType)
        {
            string workloadVersionInfo = null;
            string cpuArchitecutre = null;
            string regOutlook32Bit = null;
            string regOutlook64Bit = null;
            if (workloadType == "Outlook")
            {
                regOutlook32Bit = @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE";
                regOutlook64Bit = @"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE";

            }
            else if (workloadType == "Word")
            {
                regOutlook32Bit = @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Winword.exe";
                regOutlook64Bit = @"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\Winword.exe";
            }
            else if (workloadType == "Excel")
            {
                regOutlook32Bit = @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe";
                regOutlook64Bit = @"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\excel.exe";
            }
            else if (workloadType == "Powerpoint")
            {
                regOutlook32Bit = @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\powerpnt.exe";
                regOutlook64Bit = @"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\powerpnt.exe";
            }
            string keyPath = Registry.LocalMachine.OpenSubKey(regOutlook64Bit).GetValue("", "").ToString();

            if (string.IsNullOrEmpty(keyPath))
            {
                keyPath = Registry.LocalMachine.OpenSubKey(regOutlook32Bit).GetValue("", "").ToString();
                cpuArchitecutre = " 32 bit";

            }
            else
            {
                cpuArchitecutre = " 64 bit";
            }

            if (!string.IsNullOrEmpty(keyPath) && File.Exists(keyPath))
            {
                workloadVersionInfo = FileVersionInfo.GetVersionInfo(keyPath).FileVersion.ToString();

            }
            string officeVersion = "Microsoft Office " + workloadVersionInfo + cpuArchitecutre;
            return officeVersion;
        }

        public static LogData LogInit(string benchMarkName, object parameters, string workloadType, string resultsDirectory)
        {
            string officeSuiteVersion = GetOfficeSuiteVersion();
            var logData = new LogData();
            logData.Parameters = parameters;

            string sysPath = Path.GetFullPath(resultsDirectory);
            string outputPath = sysPath + "\\";
            logData.Benchmark = Logger.LogBenchMark(benchMarkName: benchMarkName, version: officeSuiteVersion, msofficeVersion: null);
            logData.Logging = new List<Logger.LoggingStruct>();
            logData.IterationTimings = new List<Logger.IterationTimingsStruct>();
            logData.AverageTime = new List<Logger.AverageTimeStruct>();
            logData.GeoMean = new List<Logger.GeoMeanStruct>();
            DateTime startTimeNow = DateTime.Now;
            string startTime = startTimeNow.ToString("yyyyMMdd") + "-" + startTimeNow.ToString("HHmmss");
            logData.StartTime = Logger.LogStartTime(msEpoch: "", startTime: startTime);

            DriveInfo[] allDrives = DriveInfo.GetDrives();
            Dictionary<string, string> diskInfo = new Dictionary<string, string>();

            foreach (DriveInfo d in allDrives)
            {
                if (d.IsReady == true)
                {
                    string drive_name = d.Name.Split(':')[0];
                    string available = "AvailableFreeSpaceindisk_" + drive_name + " (GB)";
                    string FreeSpace = "TotalFreeSpaceindisk_" + drive_name + " (GB)";
                    string TotalSize = "TotalSizeindisk_" + drive_name + " (GB)";
                    diskInfo.Add(available, CovertBToGB(d.AvailableFreeSpace.ToString()));
                    diskInfo.Add(FreeSpace, CovertBToGB(d.TotalFreeSpace.ToString()));
                    diskInfo.Add(TotalSize, CovertBToGB(d.TotalSize.ToString()));
                }
            }
            logData.SystemInfo = new SystemInfoStruct { };

            try
            {
                string officeVersion = OfficeInformation(workloadType);
                logData.Benchmark.msOfficeVersion = officeVersion;

                logData.SystemInfo.systemName = System.Environment.MachineName;
                logData.SystemInfo.processorModel = System.Environment.GetEnvironmentVariable("PROCESSOR_IDENTIFIER");
                logData.SystemInfo.boardInfo = MakeQuery("Manufacturer, SerialNumber, Product, Version", "Win32_BaseBoard");
                logData.SystemInfo.graphicsDriver = MakeQuery("VideoProcessor, VideoModeDescription", "Win32_VideoController");
                logData.SystemInfo.memoryConfig = MakeQuery("Capacity", "Win32_PhysicalMemory");

                logData.SystemInfo.diskConfig = diskInfo;
                logData.SystemInfo.osVersion = System.Environment.OSVersion.ToString();
                logData.SystemInfo.biosVersion = MakeQuery("Version,Name", "Win32_BIOS");
                logData.SystemInfo.batteryInfo = MakeQuery("EstimatedChargeRemaining", "Win32_Battery");
                logData.SystemInfo.memoryConfig["Capacity(GB)"] = CovertBToGB(logData.SystemInfo.memoryConfig["Capacity(GB)"]);
                if (Int16.Parse(logData.SystemInfo.batteryInfo["Charge %"]) < 4)
                    logData.Logging.Add(LogLogging(LogLevel: "Warning", TimeStamp: DateTime.Now.ToString(), Detail: "Battery is low"));

            }
            catch (Exception e)
            {
                logData.Logging.Add(Logger.LogLogging(LogLevel: "Warning",
                    TimeStamp: DateTime.Now.ToString(), Detail: $"Not able to fetch some system info : {e}"));
            }


            string jsonFileName = logData.Benchmark.name + "_" + logData.StartTime.time + ".json";
            JavaScriptSerializer ser = new JavaScriptSerializer();
            string jsonlogData = ser.Serialize(logData);
            JsonFormatter formatJson = new JsonFormatter(jsonlogData);
            string formatedJson = formatJson.Format();
            File.WriteAllText(@sysPath + "\\" + jsonFileName, formatedJson);

            return logData;
        }
        internal static ScreenResolution GetScreenDimension()
        {
            Dictionary<string, string> screenResolutionData = MakeQuery("CurrentHorizontalResolution, CurrentVerticalResolution", "Win32_VideoController");

            ScreenResolution screenResolution = new ScreenResolution();
            screenResolution.height = Convert.ToDouble(screenResolutionData["ScreenHeight"]);
            screenResolution.width = Convert.ToDouble(screenResolutionData["ScreenWidth"]);

            return screenResolution;
        }
		internal static string GetOfficeSuiteVersion()
        {
            string officeSuiteVersion = $"v{System.Reflection.Assembly.GetEntryAssembly().GetName().Version.Major.ToString()}." +
                $"{System.Reflection.Assembly.GetEntryAssembly().GetName().Version.Minor.ToString("00")}";
            return officeSuiteVersion;
        }
        public static LogData LogDeInit(LogData logData, string status, int iterations)
        {
            string message;
            if (status == "success")
            {
                message = "Program ran for " + iterations + " iterations";
            }
            else
            {
                message = "Program Failed to execute";
            }
            DateTime endTimeNow = DateTime.Now;
            string endTime = endTimeNow.ToString("yyyyMMdd") + "-" + endTimeNow.ToString("HHmmss");
            logData.StopTime = LogStopTime(msSinceEpoch: 0, Time: endTime);
            logData.Status = LogStatus(statusCode: status, detailedMessage: message);

            return logData;
        }

        public static StartTimeStruct LogStartTime(string msEpoch, string startTime)
        {
            StartTimeStruct startTimeObj = new StartTimeStruct { msSinceEpoch = msEpoch, time = startTime };
            return startTimeObj;
        }
        public static LoggingStruct LogLogging(string LogLevel, string TimeStamp, string Detail)
        {
            DateTime logTimeNow = DateTime.Now;
            string logTime = logTimeNow.ToString("yyyyMMdd") + "-" + logTimeNow.ToString("HHmmss");
            LoggingStruct loggingObj = new LoggingStruct { logLevel = LogLevel, timeStamp = logTime, detail = Detail };
            return loggingObj;
        }
        public static IterationTimingsStruct LogIterationTimings(int iter, string operationName, string value, string unit, int repetition)
        {

            IterationTimingsStruct iterationTimingsObj = new IterationTimingsStruct { iteration = iter, operationName = operationName, value = value, unit = unit, repetition = repetition };
            return iterationTimingsObj;
        }
        public static StopTimeStruct LogStopTime(int msSinceEpoch, string Time)
        {
            StopTimeStruct stopTimeObj = new StopTimeStruct { msSinceEpoch = "", time = Time };
            return stopTimeObj;
        }
        public static EtwInfoStruct LogEtwInfo(string[] eventsNames)
        {
            EtwInfoStruct etwInfoObj = new EtwInfoStruct { events = eventsNames, count = eventsNames.Length };
            return etwInfoObj;
        }
        public static StatusStruct LogStatus(string statusCode, string detailedMessage)
        {
            StatusStruct statusObj = new StatusStruct { statusCode = statusCode, detailedMessage = detailedMessage };
            return statusObj;
        }
        public static AverageTimeStruct LogAverageTime(string value, string operationName, string unit, int repetition)
        {
            AverageTimeStruct AverageTimeObj = new AverageTimeStruct { value = value, operationName = operationName, unit = unit, repetition = repetition };
            return AverageTimeObj;
        }
        public static GeoMeanStruct LogGeoMean(string value, string operationName, string unit, int repetition)
        {
            GeoMeanStruct GeoMeanObj = new GeoMeanStruct { value = value, operationName = operationName, unit = unit, repetition = repetition };
            return GeoMeanObj;
        }

        public static ExceptionStruct LogException(string exception, string message)
        {
            ExceptionStruct expObj = new ExceptionStruct { name = exception, message = message };
            return expObj;
        }
        public static ConditionsStruct LogCondition(string message)
        {
            ConditionsStruct conditionObj = new ConditionsStruct { detailedMessage = message };
            return conditionObj;
        }
        public static BehaviourStruct LogBehaviour(string message)
        {
            BehaviourStruct behaviourObj = new BehaviourStruct { detailedMessage = message };
            return behaviourObj;
        }
        class Global
        {
            public static string[] namesToBeChanged = { "Product", "EstimatedChargeRemaining", "Capacity" ,
            "CurrentHorizontalResolution", "CurrentVerticalResolution" };
            public static Dictionary<string, string> propReference = new Dictionary<string, string>(){
            { "Product", "ProductID"},
            { "EstimatedChargeRemaining", "Charge %"},
            { "Capacity", "Capacity(GB)"},
            { "CurrentHorizontalResolution", "ScreenWidth"},
            { "CurrentVerticalResolution", "ScreenHeight"}
            };

        }

        private static string ChangePropertyName(string changekey)
        {
            int pos = Array.IndexOf(Global.namesToBeChanged, changekey);

            if (pos > -1)
            {
                return Global.propReference[changekey];
            }
            else
            {
                return changekey;
            }
        }

        private static Dictionary<string, string> MakeQuery(string requiredKeywords, string infoClass)
        {
            var systemDetail = new Dictionary<string, string>();
            System.Management.ObjectQuery query = new ObjectQuery("Select " + requiredKeywords + " FROM " + infoClass);
            ManagementObjectSearcher searcher = new ManagementObjectSearcher(query);
            ManagementObjectCollection collection = searcher.Get();

            if (collection.Count > 0)
            {
                foreach (ManagementObject mo in collection)
                {
                    if (mo.Properties != null)
                    {
                        foreach (PropertyData property in mo.Properties)
                        {
                            if (!(systemDetail.ContainsKey(ChangePropertyName(property.Name))))
                            {
                                if (property.Value != null)
                                {
                                    systemDetail.Add(ChangePropertyName(property.Name), Convert.ToString(property.Value));
                                }

                            }
                        }
                    }

                }
                return systemDetail;
            }
            else
            {
                string[] requiredKeywordsList = requiredKeywords.Split(',');
                for (int property = 0; property < requiredKeywordsList.Length; property++)
                {
                    systemDetail.Add(requiredKeywordsList[property], "Not found");
                }

                return systemDetail;
            }
        }

        private static string CovertBToGB(string value)
        {
            value = Decimal.Divide(Int64.Parse(value), 1000000000).ToString();
            return value;
        }

        public static void CreateFileJson(LogData logData, Options opt)
        {
            string sysPath = Path.GetFullPath(opt.resultsDirectory);
            string jsonFileName = logData.Benchmark.name + "_" + logData.StartTime.time + ".json";

            JavaScriptSerializer ser = new JavaScriptSerializer();
            string jsonlogData = ser.Serialize(logData);
            JsonFormatter formatJson = new JsonFormatter(jsonlogData);
            string formatedJson = formatJson.Format();
            File.WriteAllText(@sysPath + "\\" + jsonFileName, formatedJson);
        }

        public static Logger.LogData Init(Options opt, string benchMakrTestName, string workloadType)
        {
            // ETW logging for Start of a program
            Utility.CallStartBenchmark();

            Utility.ValidateResultsDirectory(opt);
            if (System.Convert.ToBoolean(opt.Verbose))
            {
                Console.WriteLine($"{GetTimestamp(DateTime.Now)}: {benchMakrTestName} Test started - {GetOfficeSuiteVersion()}");
            }
                
            var logData = Logger.LogInit(benchMarkName: $"{benchMakrTestName}", parameters: opt, workloadType, opt.resultsDirectory);
            logData.Logging.Add(Logger.LogLogging(LogLevel: "Info", TimeStamp: DateTime.Now.ToString(), Detail: "Program has started"));

            return logData;

        }
        public static void ExceptionDeInit(string exceptionName, string message, Logger.LogData logData, Options opt)
        {
            logData.Exception = Logger.LogException(exception: exceptionName, message: message);
            Logger.LogDeInit(logData, "failure", 0);
            Logger.CreateFileJson(logData, opt);
            System.Environment.Exit(1);
        }
        public static String GetTimestamp(DateTime value)
        {
            return value.ToString("yyyy-MM-dd HH:mm:ssffff");
        }
        public static void LogProgramEnd(string status, Options opt)
        {
            if (System.Convert.ToBoolean(opt.Verbose))
            {
                if (status == "success")
                    Console.WriteLine($"{GetTimestamp(DateTime.Now)}: Program completed successfully ");
                else
                    Console.WriteLine($"{GetTimestamp(DateTime.Now)}: Program terminated with exception ");
            }
        }
        public static void IterationEventStart(int iter)
        {
            // ETW logging for start of iteration
            Logging.Log.Write("Start of Iteration : " + (iter + 1), new EventSourceOptions
            {
                Level = EventLevel.Verbose,
                Opcode = EventOpcode.Info
            });

        }

        public static void IterationEventEnd(int iter)
        {
            // ETW logging for end of iteration
            Logging.Log.Write("End of Iteration : " + (iter + 1), new EventSourceOptions
            {
                Level = EventLevel.Verbose,
                Opcode = EventOpcode.Info
            });

        }

        public static void IterationEventStart(string operationName, int rep, int iter)
        {
            // ETW logging for start of iteration
            string statelabel = "Start of Iteration";
            string[] operationNameEvent = operationName.Split('_');
            operationName = operationNameEvent[0] + "_" + operationNameEvent[1];

            Logging.Log.Write($"{operationName}: {String.Format("{0:0000}", rep)}-{String.Format("{0:00}", (iter + 1))} {statelabel} ", new EventSourceOptions
            {
                Level = EventLevel.Verbose,
                Opcode = EventOpcode.Info
            });

        }

        public static void IterationEventEnd(string operationName, int rep, int iter)
        {
            // ETW logging for end of iteration            
            string statelabel = "End of Iteration";
            string[] operationNameEvent = operationName.Split('_');
            operationName = operationNameEvent[0] + "_" + operationNameEvent[1];


            Logging.Log.Write($"{operationName}: {String.Format("{0:0000}", rep)}-{String.Format("{0:00}", (iter + 1))} {statelabel} ", new EventSourceOptions
            {
                Level = EventLevel.Verbose,
                Opcode = EventOpcode.Info
            });

        }

        public class Settings
        {
            public Dictionary<string, Double> averageTime = new Dictionary<string, Double>();
            public Dictionary<string, Double> GeoMean = new Dictionary<string, Double>();
            public int previousFullScreenSetting = 0;
            public Dictionary<string, List<Double>> iterationTimings = new Dictionary<string, List<Double>>();
            public List<string> fileNames = new List<string>();
            public Exception exception = null;
            public string status = null;
            public int repetition = 0;
            // getting window handle of console
            public IntPtr hWnd = Logger.GetConsoleWindow();
            public int caseID = 100;
            public Dictionary<string, List<double>> iterationTimingsCollection = new Dictionary<string, List<double>>();
            // format of timing and geomean calculation
            internal Utility.LogTimingFormat timingFormat = 0;
            // tells whether to logtiming thorugh utility
            internal int CalculateTimeCollection = 0;
            internal List<string> operationList = new List<string>();
            internal bool openFileFromInit = true;
            internal bool closeFileFromDeInit = true;
            internal ScreenResolution originalScreenSetting = new ScreenResolution();
            internal Dictionary<string, int> GeoMeanDropSetting = new Dictionary<string, int>();
        }

        public static void DeInit(Logger.LogData logData, Logger.Settings set, string status, Options opt, Exception e)
        {
            if (status == "success")
            {


                logData = Utility.LogTiming(set, opt, logData);


            }
            else
            {
                logData.Exception = Logger.LogException(exception: e.ToString(), message: e.Message);
            }

            Logger.LogDeInit(logData, status, opt.Iterations);
            Logger.CreateFileJson(logData, opt);

            // ETW logging for end of program

            Utility.CallEndBenchmark();

            // Add CSV logger information.

            CsvLogger.GenerateExcel(set, logData, opt);

            // shows up the command prompt
            //Logger.ShowWindow(set.hWnd, Logger.swShowNormal);

            /*if (status == "success")
                Console.WriteLine("Program completed successfully");
            else
                Console.WriteLine("Program terminated with exception");*/
        }

        // Json formatter
        public class StringWalker
        {
            private readonly string inputString;

            public int Index { get; private set; }
            public bool IsEscaped { get; private set; }
            public char CurrentChar { get; private set; }

            public StringWalker(string s)
            {
                inputString = s;
                this.Index = -1;
            }

            public bool MoveNext()
            {
                if (this.Index == inputString.Length - 1)
                    return false;

                if (IsEscaped == false)
                    IsEscaped = CurrentChar == '\\';
                else
                    IsEscaped = false;
                this.Index++;
                CurrentChar = inputString[Index];
                return true;
            }
        };

        public class IndentWriter
        {
            private readonly StringBuilder result = new StringBuilder();
            private int indentLevel;

            public void Indent()
            {
                indentLevel++;
            }

            public void UnIndent()
            {
                if (indentLevel > 0)
                    indentLevel--;
            }

            public void WriteLine(string line)
            {
                result.AppendLine(CreateIndent() + line);
            }

            private string CreateIndent()
            {
                StringBuilder indent = new StringBuilder();
                for (int i = 0; i < indentLevel; i++)
                    indent.Append("    ");
                return indent.ToString();
            }

            public override string ToString()
            {
                return result.ToString();
            }
        };

        public class JsonFormatter
        {
            private readonly StringWalker walker;
            private readonly IndentWriter writer = new IndentWriter();
            private readonly StringBuilder currentLine = new StringBuilder();
            private bool quoted;

            public JsonFormatter(string json)
            {
                walker = new StringWalker(json);
                ResetLine();
            }

            public void ResetLine()
            {
                currentLine.Length = 0;
            }

            public string Format()
            {
                while (MoveNextChar())
                {
                    if (this.quoted == false && this.IsOpenBracket())
                    {
                        this.WriteCurrentLine();
                        this.AddCharToLine();
                        this.WriteCurrentLine();
                        writer.Indent();
                    }
                    else if (this.quoted == false && this.IsCloseBracket())
                    {
                        this.WriteCurrentLine();
                        writer.UnIndent();
                        this.AddCharToLine();
                    }
                    else if (this.quoted == false && this.IsColon())
                    {
                        this.AddCharToLine();
                        this.WriteCurrentLine();
                    }
                    else
                    {
                        AddCharToLine();
                    }
                }
                this.WriteCurrentLine();
                return writer.ToString();
            }

            private bool MoveNextChar()
            {
                bool success = walker.MoveNext();
                if (this.IsApostrophe())
                {
                    this.quoted = !quoted;
                }
                return success;
            }

            public bool IsApostrophe()
            {
                return this.walker.CurrentChar == '"' && this.walker.IsEscaped == false;
            }

            public bool IsOpenBracket()
            {
                return this.walker.CurrentChar == '{'
                    || this.walker.CurrentChar == '[';
            }

            public bool IsCloseBracket()
            {
                return this.walker.CurrentChar == '}'
                    || this.walker.CurrentChar == ']';
            }

            public bool IsColon()
            {
                return this.walker.CurrentChar == ',';
            }

            private void AddCharToLine()
            {
                this.currentLine.Append(walker.CurrentChar);
            }

            private void WriteCurrentLine()
            {
                string line = this.currentLine.ToString().Trim();
                if (line.Length > 0)
                {
                    writer.WriteLine(line);
                }
                this.ResetLine();
            }
        };

    }
}