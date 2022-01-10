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
        [Option('b', "ToAddress", Required = false, HelpText = "Address to send the mail (Specify in a single line with commas)", Default = "Sujith L,49th Cross,Bangalore,Karnataka.Krish S,Greenways,Mumbai,Maharastra")]
        public string ToAddress { get; set; }

        [Option('a', "FromAddress", Required = false, HelpText = "Return/From Address (Specify in a single line with commas and while entering multiple addresses separate with .)", Default = "Ajay K, 1st Street, Chennai, Tamil Nadu.John L, 2nd St, Hyderabad, Telangana")]
        public string FromAddress { get; set; }

        public const bool inputFileFlag = true;
        public const bool outputFileFlag = true;
    }

    class SpecificArguments
    {
        // Command line option for InsertEnvelope
        public static Options ParseArgument(string[] args, Logger.Settings set)
        {
            Options options = new Options();

            if (args.Length != 0)
            {
                Utility.SetCaseID(args, set);

                // Invoke InsertEnvelope default
                if (args[0] == "default")
                {
                    options.InputFileName = "..\\input\\Monte_Cristo_original_small.docx";
                    options.OutputFileName = "InsertEnvelopeOutput_Default.docx";
                    options.Iterations = 2;
                    options.FromAddress = "Ajay K,1st Street,Chennai,Tamil Nadu.John L,2nd St,Hyderabad,Telangana";
                    options.ToAddress = "Sujith L,49th Cross,Bangalore,Karnataka.Krish S,Greenways,Mumbai,Maharastra";
                }
            }
            Arguments.ParseArgument(args, ref options);
            return options;
        }
    }

    static class InsertEnvelope
    {
        static void CreateCSV(Options opt, Logger.LogData logData, string status)
        {
            string operationname = logData.Benchmark.name;
            string startTime = logData.StartTime.time;
            string sysPath = Path.GetFullPath(opt.resultsDirectory);
            string csvfile = sysPath + "\\" + operationname + "_" + startTime + ".csv";

            string[] lines = File.ReadAllLines(csvfile);
            string[] FromAddress = opt.FromAddress.Split('.');
            string[] ToAddress = opt.ToAddress.Split('.');

            if (lines.Length == 0)
            {
                status = "failure";
            }

            //add new column to the header row
            lines[0] += ",InputFileName,FromAddress,ToAddress";

            //add new column value for each row.
            for (int i = 1; i < lines.Length; i++)
            {
                lines[i] += "," + Path.GetFileName(opt.InputFileName) + "," + FromAddress[i - 1].Replace(",", "/") 
                    + "," + ToAddress[i - 1].Replace(",", "/");
            }

            //write the new content
            File.WriteAllLines(csvfile, lines);
            Logger.LogProgramEnd(status, opt);
        }

        static void ValidateIterationValues(out string[] fromaddress, out string[] toaddress, Options opt, Logger.LogData logData)
        {
            fromaddress = opt.FromAddress.Split('.');
            toaddress = opt.ToAddress.Split('.');
            
            if(fromaddress.Length != toaddress.Length)
            {
                Console.WriteLine("Length among fromaddress and toaddress doesn't match with each other. Program failed");
                Logger.ExceptionDeInit("Length Mismatch Error", "Length among fromaddress and toaddress doesn't match with each other", logData, opt);
            }

            if(opt.Iterations == -1)
            {
                opt.Iterations = fromaddress.Length;
            }

            if(opt.Iterations >= fromaddress.Length)
            {
                opt.FromAddress = "";
                opt.ToAddress = "";
                for(int i = 0; i < opt.Iterations; i++)
                {
                    opt.FromAddress += fromaddress[i % fromaddress.Length] + ".";
                    opt.ToAddress += toaddress[i % toaddress.Length] + ".";
                }

                string[] fromaddresslist = new string[opt.Iterations];
                string[] toaddresslist = new string[opt.Iterations];

                for (int i = 0; i < opt.Iterations; i++)
                {
                    fromaddresslist[i] = "From:\n" + fromaddress[i % fromaddress.Length];
                    fromaddresslist[i] = fromaddresslist[i].Replace(",", ",\n");
                    toaddresslist[i] = "To:\n" + toaddress[i % toaddress.Length];
                    toaddresslist[i] = toaddresslist[i].Replace(",", ",\n");
                }
                fromaddress = fromaddresslist;
                toaddress = toaddresslist;
            }
            else if (opt.Iterations < fromaddress.Length)
            {

                opt.FromAddress = "";
                opt.ToAddress = "";
                for (int i = 0; i < opt.Iterations; i++)
                {
                    opt.FromAddress += fromaddress[i % fromaddress.Length] + ".";
                    opt.ToAddress += toaddress[i % toaddress.Length] + ".";
                }

                string[] fromaddresslist = new string[opt.Iterations];
                string[] toaddresslist = new string[opt.Iterations];

                for (int i = 0; i < opt.Iterations; i++)
                {
                    fromaddresslist[i] = "From: \n" + fromaddress[i];
                    fromaddresslist[i] = fromaddresslist[i].Replace(",", ",\n");
                    toaddresslist[i] = "To: \n" + toaddress[i];
                    toaddresslist[i] = toaddresslist[i].Replace(",", ",\n");
                }
                fromaddress = fromaddresslist;
                toaddress = toaddresslist;
            }
        }

        static void WordInsertEnvelope(Word.Application app, List<Word._Document> documents, List<Word.Window> docWindows, Logger.Settings set,
            Options opt, Logger.LogData logData, int repeat)
        {
            for (int rep = 1; rep <= repeat; rep++)
            {
                string operationName = "InsertEnvelope_" + rep.ToString();
                set.repetition = rep;

                string[] FromAddress = null;
                string[] ToAddress = null;

                set.fileNames.Add(opt.InputFileName);
                ValidateIterationValues(out FromAddress, out ToAddress, opt, logData);
                Utility.ValidateOutputFiles(rep, opt, logData);

                try
                {
                    Utility.WordInit(out app, out documents, out docWindows, operationName, set, opt, logData);
                    set.iterationTimings.Add(operationName, new List<double>());

                    Thread.Sleep(opt.SeparationPause);

                    for (int iter = 0; iter < opt.Iterations; iter++)
                    {
                        Stopwatch watch = new Stopwatch();

                        documents[0].UndoClear();
                        Word.UndoRecord ur = app.UndoRecord;
                        ur.StartCustomRecord($"Inserting Envelope {iter}");

                        Thread.Sleep(opt.SeparationPause);
                        Logger.IterationEventStart(operationName, rep, iter);

                        //Timer Starts
                        watch.Start();

                        //Insert the Envelope with specified addresses
                        documents[0].Envelope.Insert(Address: ToAddress[iter], ReturnAddress: FromAddress[iter]);

                        //Timer Stops
                        watch.Stop();
                        Logger.IterationEventEnd(operationName, rep, iter);

                        Thread.Sleep(opt.IterationPause);

                        ur.EndCustomRecord();
                        documents[0].Undo();

                        set.iterationTimings[operationName].Add(watch.ElapsedMilliseconds / 1000.0);
                        Utility.IterationEnd(operationName, iter, logData, set, opt.IterationPause);
                    }

                    // Save the document
                    documents[0].SaveAs(Utility.GetFileName(opt.OutputFileName, rep, opt.Iterations));
                    Thread.Sleep(opt.SeparationPause);

                    Utility.WordDeInit(app, documents, set, opt.SeparationPause);
                    set.status = "success";
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

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
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

            string benchMarkTestName = "Word_InsertEnvelope";
            Logger.LogData logData = Logger.Init(opt, benchMarkTestName, "Word");

            Utility.ValidateInputFiles(opt, logData);

            // Decalre the Word objects
            Word.Application app = null;
            List<Word._Document> documents = null;
            List<Word.Window> docWindows = null;

            Console.WriteLine("Inserting Envelope !");
            WordInsertEnvelope(app, documents, docWindows, set, opt, logData, (opt.runs));

        }
    }
}