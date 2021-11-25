using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Diagnostics;
using System.IO;
using System.Text;

using Excel = Microsoft.Office.Interop.Excel;

using Helper;
using CommandLine;

namespace Workload
{
    public class Options : CommonOptions
    {
        [Option("SheetNumber", Required = false, HelpText = "Sheet number eg: 1", Default = 1)]
        public int SheetNumber { get; set; }

        [Option("RangeList", Required = false, HelpText = "Cell range eg: A1:B400,C1:D400", Default = "A1:A400,B1:B400,C1:C400,D1:D400,E1:E400," +
            "F1:F400,G1:G400,H1:H400,I1:I400,J1:J400")]
        public string RangeList { get; set; }

        [Option("AlignmentTypeList", Required = false, HelpText = "AlignmentTypeList for each iteration eg:HAlign,HAlign ", Default = "HAlign,HAlign,HAlign,HAlign," +
            "HAlign,VAlign,VAlign,VAlign,VAlign,VAlign")]
        public string AlignmentTypeList { get; set; }

        [Option("SubAlignmentTypeList", Required = false, HelpText = "SubAlignmentList for each iteration eg:Right,Left", Default = "Right,Left,Justify,Distributed," +
            "Center,Top,Justify,Distributed,Center,Bottom")]
        public string SubAlignmentTypeList { get; set; }

        public const bool inputFileFlag = true;
        public const bool outputFileFlag = true;
    }

    static class SpecificArguments
    {
        // Command line option for TextAlignment
        public static Options ParseArgument(string[] args, Logger.Settings set)
        {
            Options options = new Options();

            if (args.Length != 0)
            {
                Utility.SetCaseID(args, set);

                // Invoke TextAlignment default
                if (args[0] == "default")
                {
                    options.InputFileName = "..\\input\\VoterAnalysis-TextAlignment.xlsm";
                    options.OutputFileName = "ResultTextAlignment.xlsm";
                    options.Iterations = 3;
                    options.SheetNumber = 1;
                    options.RangeList = "A1:A400;B1:B400;C1:C400;D1:D400,E1:E400;F1:F400;G1:G400;H1:H400,I1:I400;J1:J400"; // RangeList
                    options.AlignmentTypeList = "HAlign;HAlign;HAlign;HAlign,HAlign;VAlign;VAlign;VAlign,VAlign;VAlign"; // HAlign or VAlign                    
                    options.SubAlignmentTypeList = "Right;Left;Justify;Distributed,Center;Top;Justify;Distributed,Center;Bottom";
                    // subAlignmenetsHAlign = { "Right","Left","Justify","Distributed","Center","General","Fill","CenterAcrossSelection" }
                    // subAlignmenetsVAlign = { "Top","Justify","Distributed","Center","Bottom" };
                }
            }
            Arguments.ParseArgument(args, ref options);
            return options;
        }
    }

    static class TextAlignment
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
            string[] RangeList = opt.RangeList.Split(',');
            string[] AlignmentList = opt.AlignmentTypeList.Split(',');
            string[] SubAlignmentList = opt.SubAlignmentTypeList.Split(',');

            //add new column to the header row
            lines[0] += ",InputFileName,SheetNumber,Range,Alignment,SubAlignment";
            
            //add new column value for each row.
            for (int i = 1; i < lines.Length; i++)
            {
                lines[i] += "," + Path.GetFileName(opt.InputFileName) + "," + opt.SheetNumber + "," + RangeList[i-1] + 
                    "," + AlignmentList[i-1] + "," + SubAlignmentList[i-1];
            }
            
            //write the new content
            File.WriteAllLines(csvfile, lines);
            Logger.LogProgramEnd(status,opt);
        }

        class AlignmentSettings
        {
            public string[] alignments = { "HAlign", "VAlgin" };
            public string[] subAlignmenetsHAlign = { "Right","Left","Justify","Distributed",
                "Center","General","Fill","CenterAcrossSelection" };
            public string[] subAlignmenetsVAlign = { "Top","Justify","Distributed",
                "Center","Bottom" };

            public Dictionary<string, Excel.XlHAlign> hAlignDict = new Dictionary<string, Excel.XlHAlign> {
                {"Right",Excel.XlHAlign.xlHAlignRight },
                {"Left",Excel.XlHAlign.xlHAlignLeft },
                {"Justify",Excel.XlHAlign.xlHAlignJustify },
                {"Distributed",Excel.XlHAlign.xlHAlignDistributed },
                {"Center",Excel.XlHAlign.xlHAlignCenter },
                {"General",Excel.XlHAlign.xlHAlignGeneral },
                {"Fill",Excel.XlHAlign.xlHAlignFill },
                {"CenterAcrossSelection",Excel.XlHAlign.xlHAlignCenterAcrossSelection }

            };

            public Dictionary<string, Excel.XlVAlign> vAlignDict = new Dictionary<string, Excel.XlVAlign> {
                {"Top",Excel.XlVAlign.xlVAlignTop },
                {"Justify",Excel.XlVAlign.xlVAlignJustify },
                {"Distributed",Excel.XlVAlign.xlVAlignDistributed },
                {"Center",Excel.XlVAlign.xlVAlignCenter },
                {"Bottom",Excel.XlVAlign.xlVAlignBottom }


            };
        }


        // Excel TextAlignment Validations
        static void ValidateList(out string[] rangeList ,out string[] alignmentList ,out string[] subAlignmentList, Logger.LogData logData, Options opt)
        {
            rangeList = opt.RangeList.Split(',');
            alignmentList = opt.AlignmentTypeList.Split(',');
            subAlignmentList = opt.SubAlignmentTypeList.Split(',');

            // Check the total number of iterations and columns number is same
            if ((alignmentList.Length != rangeList.Length) || (subAlignmentList.Length != rangeList.Length))
            {
                Console.WriteLine("Length of RangeList, AlignmentTypeList, SubAlignmentTypeList doesnt match among each other. Program failed");
                Logger.ExceptionDeInit("Invalid Input", "Length of RangeList, AlignmentTypeList, SubAlignmentTypeList doesnt match among each other", logData,opt);
            }

            // Check the total number of iterations and columns number 
            if (opt.Iterations == -1)
            {
                opt.Iterations = rangeList.Length;
            }
            else if (opt.Iterations > rangeList.Length)
            {
                string[] modRangeList = new string[opt.Iterations];
                string[] modAlignmentTypeList = new string[opt.Iterations];
                string[] modSubAlignmentTypeList = new string[opt.Iterations];
                for (int i = 0; i < opt.Iterations; i++)
                {
                    modRangeList[i] = rangeList[(i % rangeList.Length)];
                    modAlignmentTypeList[i] = alignmentList[(i % alignmentList.Length)];
                    modSubAlignmentTypeList[i] = subAlignmentList[(i % subAlignmentList.Length)];
                }
                rangeList = modRangeList;
                alignmentList = modAlignmentTypeList;
                subAlignmentList = modSubAlignmentTypeList;
                opt.RangeList = string.Join(",", rangeList);
                opt.AlignmentTypeList = string.Join(",", alignmentList);
                opt.SubAlignmentTypeList = string.Join(",", subAlignmentList);
            }
            else if (opt.Iterations < rangeList.Length)
            {
                string[] modRangeList = new string[opt.Iterations];
                string[] modAlignmentTypeList = new string[opt.Iterations];
                string[] modSubAlignmentTypeList = new string[opt.Iterations];
                for (int i = 0; i < opt.Iterations; i++)
                {
                    modRangeList[i] = rangeList[i];
                    modAlignmentTypeList[i] = alignmentList[i];
                    modSubAlignmentTypeList[i] = subAlignmentList[i];
                }
                rangeList = modRangeList;
                alignmentList = modAlignmentTypeList;
                subAlignmentList = modSubAlignmentTypeList;
                opt.RangeList = string.Join(",", rangeList);
                opt.AlignmentTypeList = string.Join(",", alignmentList);
                opt.SubAlignmentTypeList = string.Join(",", subAlignmentList);
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

        
        static void ValidateInput(string[] alignmentList, string[] subAlignmentList, AlignmentSettings alignSet, Logger.LogData logData,Options opt)
        {
           
            for (int iter=0; iter< alignmentList.Length; iter++)
            {
                if (alignmentList[iter] == "HAlign")
                {
                    if (!(alignSet.subAlignmenetsHAlign.Contains(subAlignmentList[iter])))
                    {
                        Console.WriteLine("Invalid SubAlignment arguments. Program failed");
                        Logger.ExceptionDeInit("Invalid Input", "Invalid SubAlignment arguments", logData,opt);
                    }
                    
                }
                else if (alignmentList[iter] == "VAlign")
                {
                    if (!(alignSet.subAlignmenetsVAlign.Contains(subAlignmentList[iter])))
                    {
                        Console.WriteLine("Invalid SubAlignment arguments. Program failed");
                        Logger.ExceptionDeInit("Invalid Input", "Invalid SubAlignment arguments", logData,opt);
                    }
                    
                }
                else
                {
                    Console.WriteLine("Invalid Alginment argument. Program failed");
                    Logger.ExceptionDeInit("Invalid Input", "Invalid Alginment argument", logData,opt);
                }

            }

        }

        


        static void ExcelTextAlignment(Excel.Application app, List<Excel._Workbook> workbooks, List<Excel._Worksheet> worksheets, Logger.Settings set,
            Options opt, Logger.LogData logData, int repeat)
        {
            for (int rep = 1; rep <= repeat; rep++)
            {
                string operationName = "TextAlignment_" + rep.ToString();
                set.repetition = rep;
                string[] rangeList = null;
                string[] alignmentList = null;
                string[] subAlignmentList = null;

                ValidateList(out rangeList, out alignmentList, out subAlignmentList, logData, opt);

                AlignmentSettings alignSet = new AlignmentSettings();
                Utility.ValidateOutputFiles(rep, opt, logData);
                //ValidateInput(alignmentList, subAlignmentList, alignSet, logData);
                Thread.Sleep(opt.SeparationPause);

                
                set.fileNames.Add(opt.InputFileName);
                set.iterationTimings.Add(operationName, new List<double>());

                try
                {
                    Utility.ExcelInit(out app, out workbooks, out worksheets, operationName, set, opt);

                    ValidateSheetNumber(workbooks, worksheets, logData, opt.SheetNumber);
                    // Adding worksheet to worksheet list and activating
                    worksheets.Add(workbooks[0].Sheets[opt.SheetNumber]);
                    worksheets[0].Activate();

                    Thread.Sleep(opt.SeparationPause);

                    for (int iter = 0; iter < opt.Iterations; iter++)
                    {
                        string[] rangeIterationList = rangeList[iter].Split(';');
                        string[] alignmentIterationList = alignmentList[iter].Split(';');
                        string[] subAlignmentIterationList = subAlignmentList[iter].Split(';');
                        if ((alignmentIterationList.Length != rangeIterationList.Length) || (subAlignmentIterationList.Length != rangeIterationList.Length))
                        {
                            Console.WriteLine("Length of RangeList, AlignmentTypeList, SubAlignmentTypeList for iteration doesnt match among each other. Program failed");
                            Logger.ExceptionDeInit("Invalid Input", "Length of RangeList, AlignmentTypeList, SubAlignmentTypeList doesnt match among each other", logData,opt);
                        }
                        ValidateInput(alignmentIterationList, subAlignmentIterationList, alignSet, logData,opt);
                    }
                    for (int iter = 0; iter < opt.Iterations; iter++)
                    {
                        string[] rangeIterationList = rangeList[iter].Split(';');
                        string[] alignmentIterationList = alignmentList[iter].Split(';');
                        string[] subAlignmentIterationList = subAlignmentList[iter].Split(';');

                        // Measure timings
                        Stopwatch watch = new Stopwatch();

                        for (int nestediter = 0; nestediter < rangeIterationList.Length; nestediter++)
                        {
                            Excel.Range range = worksheets[0].Range[rangeIterationList[nestediter]];

                            // Select the entire range
                            range.Select();
                            range.Activate();

                            Excel.XlHAlign horizontalAlignmentType = 0;
                            Excel.XlVAlign verticalAlignmentType = 0;

                            if (alignmentIterationList[nestediter] == "HAlign")
                            {
                                horizontalAlignmentType = alignSet.hAlignDict[subAlignmentIterationList[nestediter]];
                            }
                            else if (alignmentIterationList[nestediter] == "VAlign")
                            {
                                verticalAlignmentType = alignSet.vAlignDict[subAlignmentIterationList[nestediter]];
                            }

                            if (nestediter == 0)
                            {
                                Logger.IterationEventStart(operationName, rep, iter);
                            }

                            watch.Start();

                            if (alignmentIterationList[nestediter] == "HAlign")
                            {
                                range.HorizontalAlignment = horizontalAlignmentType;
                            }
                            else if (alignmentIterationList[nestediter] == "VAlign")
                            {
                                range.VerticalAlignment = verticalAlignmentType;
                            }

                            watch.Stop();
                        }

                        Logger.IterationEventEnd(operationName, rep, iter);

                        set.iterationTimings[operationName].Add(watch.ElapsedMilliseconds / 1000.0);
                        Utility.IterationEnd(operationName, iter, logData, set, opt.IterationPause);

                    }

                    Thread.Sleep(opt.SeparationPause);
                    // Save the workbook.
                    workbooks[0].SaveAs(Utility.GetFileName(opt.OutputFileName, rep, opt.Iterations));

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
            CreateCSV(opt, logData);
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

            string benchMarkTestName = "Excel_TextAlignment";
            Logger.LogData logData = Logger.Init(opt, benchMarkTestName, "Excel");

            Utility.ValidateInputFiles(opt, logData);
            // Decalre the Excel objects
            Excel.Application app = null;
            List<Excel._Workbook> workbooks = null;
            List<Excel._Worksheet> worksheets = null;

            
            ExcelTextAlignment(app, workbooks, worksheets, set, opt, logData, (opt.runs));

        }

    }
}
