using System;
using System.IO;
using System.Text;
using System.Linq;
using System.Collections.Generic;
using Workload; 

namespace Helper
{
    class CsvLogger
    {
            
        /// Creates a CSV file from the input provided        
        public static void GenerateExcel(Logger.Settings set, Logger.LogData logData, Options opt)
        {
            string operationName = logData.Benchmark.name;
            string startTime = logData.StartTime.time;
            string sysPath = Path.GetFullPath(opt.resultsDirectory);
            string excelFileName = sysPath + "\\" + operationName + "_" + startTime + ".csv";
                        
            System.Text.StringBuilder csv = new StringBuilder(); 
            var writeLine = "";
            Dictionary<string, List<Double>> timing = set.iterationTimingsCollection;
            string workloadVersion =  Logger.GetOfficeSuiteVersion();
            writeLine = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13}", "Case_ID", "Test", "Operation",
                "Iterations", "Runs",
                "Avg", "Min", "0.25", "0.5", "0.75", "Max", "ScreenSetting", "ScreenSize(WxH)", "TestSuiteVersion");
            csv.AppendLine(writeLine);
            string screenSetting = "";
            string screenSize = "";
            if (opt.Display == 1)
            {
                screenSetting = "FullScreen";
                screenSize = "N/A";
            }
            else
            {
                screenSetting = "NormalUserDefined";
                screenSize = $"{opt.DisplayWidth}x{opt.DisplayHeight}";
            }
            int iterCounter = 1;

            foreach (var iterationname in timing.Keys)
            {
                // sorting to calculate
                Tuple<double, double, double> quartileValues = Utility.Quartiles(set.iterationTimingsCollection[iterationname].OrderBy(o => o).ToArray());

                string[] iteration = iterationname.Split('_');
                int iterationNo = Convert.ToInt32(iteration[iteration.Length - 1]);

                writeLine = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13}", set.caseID + iterCounter,
                    logData.Benchmark.name,
                    set.operationList[iterCounter - 1], iterCounter, opt.runs, set.iterationTimingsCollection[iterationname].Average().ToString("0.000"),
                    set.iterationTimingsCollection[iterationname].Min(),
                    quartileValues.Item1,
                    quartileValues.Item2,
                    quartileValues.Item3,
                    set.iterationTimingsCollection[iterationname].Max(), screenSetting, screenSize, workloadVersion);
                iterCounter += 1;
                csv.AppendLine(writeLine);

            }

            File.WriteAllText(excelFileName, csv.ToString());

        }
    }
}