using System;
using System.Collections.Generic;
using System.Threading;
using System.Diagnostics;
using System.IO;

using CommandLine;
using Word = Microsoft.Office.Interop.Word;

using Helper;

namespace Workload
{
    public class Options : CommonOptions
    {
        [Option("LoadIterationList", Required = false, HelpText = "Times on which each document should be loaded ", Default = "1,1,1,1,1,5")]
        public string LoadIterationList { get; set; }

        [Option('z', "ZoomPercentage", Required = false, HelpText = "Zoom percentage", Default = 100)]
        public int ZoomPercentage { get; set; }

        public const bool inputFileFlag = true;
        public const bool outputFileFlag = true;
    }

    class SpecificArguments
    {
        // Command line option for JenkaLoad
        public static Options ParseArgument(string[] args, Logger.Settings set)
        {
            Options options = new Options();

            if (args.Length != 0)
            {
                Utility.SetCaseID(args, set);

                // Invoke JenkaLoad default
                if (args[0] == "default")
                {
                    options.InputFileName = "..\\input\\source_document.docx,..\\input\\destination_document.docx,..\\input\\Solar_panel_modified.docx,..\\input\\Solar_panel_saved.docx,..\\input\\Solar_panel.docx";
                    options.OutputFileName = "JenkaLoadOutput_Default.docx";
                    options.LoadIterationList = "3,3,1,1,1";
                    options.Iterations = 5;
                    options.ZoomPercentage = 100;
                }
            }
            Arguments.ParseArgument(args, ref options);
            return options;
        }
    }

    static class JenkaLoad
    {
        static void CreateCSV(Options opt, Logger.LogData logData, string status)
        {
            string operationname = logData.Benchmark.name;
            string startTime = logData.StartTime.time;
            string sysPath = Directory.GetCurrentDirectory();
            string csvfile = sysPath + "\\" + operationname + "_" + startTime + ".csv";

            string[] lines = File.ReadAllLines(csvfile);
            string[] InputFileName = opt.InputFileName.Split(',');

            if (lines.Length == 0)
            {
                throw new InvalidOperationException("The file is empty");
            }

            //add new column to the header row
            lines[0] += ",InputFileName,ZoomPercentage";

            //add new column value for each row.
            for (int i = 1; i < lines.Length; i++)
            {
                lines[i] += "," + Path.GetFileName(InputFileName[i - 1]) + "," + opt.ZoomPercentage ;
            }

            //write the new content
            File.WriteAllLines(csvfile, lines);
            Logger.LogProgramEnd(status, opt);
        }


        static void ValidateIterationValues(out string[] InputFileName, out string[] loadIterationList, Options opt, Logger.LogData logData)
        {
            InputFileName = opt.InputFileName.Split(',');
            loadIterationList = opt.LoadIterationList.Split(',');

            int total = 0;

            for (int iter = 0; iter < loadIterationList.Length; iter++)
            {
                total += int.Parse(loadIterationList[iter]);
            }

            if (opt.Iterations == InputFileName.Length && InputFileName.Length == total)
            {
                return;
            }

            int max = InputFileName.Length > loadIterationList.Length ? InputFileName.Length : loadIterationList.Length;

            if (opt.Iterations == -1 && InputFileName.Length == loadIterationList.Length)
            {
                opt.Iterations = max;
                return;
            }
            else if (opt.Iterations == -1 || opt.Iterations < max)
            {
                opt.Iterations = max;
            }

            string[] inputlist = new string[opt.Iterations];
            string[] iterationlist = new string[opt.Iterations];

            for (int i = 0; i < opt.Iterations; i++)
            {
                inputlist[i] = InputFileName[(i % InputFileName.Length)];
                iterationlist[i] = loadIterationList[(i % loadIterationList.Length)];
            }
            loadIterationList = iterationlist;
            opt.LoadIterationList = string.Join(",", loadIterationList);
            InputFileName = inputlist;
            opt.InputFileName = string.Join(",", InputFileName);
        }

        static string[] ValidateIterationCount(string[] inputfilename, int[] loadIterationList, Options opt, Logger.LogData logData)
        {
            int total = 0;

            for (int iter = 0; iter < loadIterationList.Length; iter++)
            {
                total += loadIterationList[iter];
            }

            if (opt.Iterations == inputfilename.Length && inputfilename.Length == total)
            {
                return inputfilename;
            }

            string[] inputfilenamelist = new string[total];
            opt.Iterations = total;
            int count = 0;
            for (int iter = 0; iter < inputfilename.Length; iter++)
            {
                for (int j = 0; j < loadIterationList[iter]; j++)
                {
                    inputfilenamelist[count] = inputfilename[iter];
                    count++;
                }
            }
            inputfilename = inputfilenamelist;
            opt.InputFileName = string.Join(",", inputfilename);
            return inputfilename;
        }

        // Load Document
        static void WordJenkaLoad(Word.Application app, List<Word._Document> documents, List<Word.Window> docWindows, Logger.Settings set,
            Options opt, Logger.LogData logData, int repeat)
        {
            for (int rep = 1; rep <= repeat; rep++)
            {
                string operationName = "JenkaLoad_" + rep.ToString();
                set.repetition = rep;

                string[] loaditerationList = null;
                string[] InputFileName = null;
                ValidateIterationValues(out InputFileName, out loaditerationList, opt, logData);
                Utility.ValidateOutputFiles(rep, opt, logData);

                try
                {
                    set.openFileFromInit = false;
                    set.iterationTimings.Add(operationName, new List<double>());

                    int[] loaditerations = Array.ConvertAll(loaditerationList, int.Parse);
                    InputFileName = ValidateIterationCount(InputFileName, loaditerations, opt, logData);
                    Utility.ValidateOutputFiles(rep, opt, logData);
                    Thread.Sleep(opt.SeparationPause);

                    for (int iter = 0; iter < opt.Iterations; iter++)
                    {
                        Utility.WordInit(out app, out documents, out docWindows, operationName, set, opt, logData);
                        var watch = new System.Diagnostics.Stopwatch();
                        set.fileNames.Add(InputFileName[iter]);

                        if(iter == 0 || loaditerations.Length > 1)
                        {
                            int flag = 1;
                            if (iter != 0 || iter != loaditerations[1])
                                flag = 0;

                            if (flag == 1)
                            {
                                Utility.StartFileOpen(operationName, set.repetition, opt.Iterations, $"_drop_{iter}");
                                Word.Document doc = app.Documents.Open(InputFileName[iter]);
                                Utility.StopFileOpen(operationName, set.repetition, opt.Iterations, $"_drop_{iter}");

                                doc.Close();
                            }
                            Thread.Sleep(opt.SeparationPause);
                        }

                        Utility.StartFileOpen(operationName, set.repetition, opt.Iterations, $"Word Start File Open {(iter + 1)}_{set.repetition}");
                        watch.Start();

                        documents.Add(app.Documents.Open(InputFileName[iter]));
                        
                        watch.Stop();
                        // ETW logging for opening a file            
                        Utility.StopFileOpen(operationName, set.repetition, opt.Iterations, $"Word Stop File Open {(iter + 1)}_{set.repetition}");

                        Thread.Sleep(opt.IterationPause);

                        documents[0].Activate();
                        docWindows = new List<Word.Window>();

                        foreach (Word.Window window in app.Windows)
                        {
                            if (window.Document == documents[0])
                            {
                                docWindows.Add(window);
                            }
                        }

                        if (docWindows[0] == null)
                        {
                            Console.WriteLine("Cannot get a reference to the destination document window!");
                            Logger.ExceptionDeInit("Exception: Destination document window not found", "Cannot get a reference to the destination document window!", logData, opt);
                        }
                        docWindows[0].Panes[1].View.Zoom.Percentage = opt.ZoomPercentage;
                        docWindows[0].View.FullScreen = true;

                        set.iterationTimings[operationName].Add(watch.ElapsedMilliseconds / 1000.0);
                        Utility.IterationEnd(operationName, iter, logData, set, opt.IterationPause);

                        documents[0].SaveAs(Utility.GetFileName(opt.OutputFileName, rep, iter, InputFileName[iter]));
                        Thread.Sleep(opt.SeparationPause);
                        Utility.WordDeInit(app, documents, set, opt.SeparationPause);
                    }
                    set.status = "success";
                    Thread.Sleep(opt.SeparationPause);

                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception raised. Program failed");
                    Console.WriteLine(e);
                    Console.WriteLine(e.Message);

                    Utility.WordDeInit(app, documents, set, opt.SeparationPause);
                    set.exception = e;
                    set.status = "failure";
                    Thread.Sleep(opt.SeparationPause);
                    return;
                }
                app = null;
                documents = null;
                docWindows = null;
                set.fileNames = new List<string>();
            }
            Thread.Sleep(opt.SeparationPause);
            Logger.DeInit(logData, set, set.status, opt, set.exception);
            CreateCSV(opt, logData, set.status);
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

            string benchMarkTestName = "Word_JenkaLoad";
            Logger.LogData logData = Logger.Init(opt, benchMarkTestName, "Word");

            Utility.ValidateInputFiles(opt, logData);

            // Decalre the Word objects
            Word.Application app = null;
            List<Word._Document> documents = null;
            List<Word.Window> docWindows = null;

            WordJenkaLoad(app, documents, docWindows, set, opt, logData, (opt.runs));
        }
    }
}
