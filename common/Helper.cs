using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CommandLine;
using Workload;

namespace Helper
{
    

    //Common options for all the workloads
    public class CommonOptions
    {
        
        [Option('i', "InputFileName", Required = Options.inputFileFlag, HelpText = "Absolute path of Input filename or Relative path (In Current directory only)")]
        public string InputFileName { get; set; }
        
        [Option('o', "OutputFileName", Required = Options.outputFileFlag, HelpText = "Absolute path of Output filename or Relative path (In Current directory only)")]
        public string OutputFileName { get; set; }

        [Option('n', "Iterations", Required = false, HelpText = "Number of iterations", Default = -1)]
        public int Iterations { get; set; }

        [Option('p', "IterationPause", Required = false, HelpText = "Pause between iterations", Default = 1000)]
        public int IterationPause { get; set; }

        [Option('s', "SeparationPause", Required = false, HelpText = "Sleep time between operations", Default = 2000)]
        public int SeparationPause { get; set; }

        [Option("StartupPause", Required = false, HelpText = "Sleep time between application start and workload execution", Default = 2000)]
        public int StartupPause { get; set; }
        
        [Option("Display", Required = false, HelpText = "Display fullscreen setting: 1 (Full Screen) or 2 (custom window)", Default = 2)]
        public int Display { get; set; }

        [Option("DisplayHeight", Required = false, HelpText = "Display screen height ", Default = 700)]
        public double DisplayHeight { get; set; }

        [Option("DisplayWidth", Required = false, HelpText = "Display screen width ", Default = 1200)]
        public double DisplayWidth { get; set; }

        [Option('r', "runs", Required = false, HelpText = "Number to times run the application", Default = 1)]
        public int runs { get; set; }

        [Option('v', "verbose", Required = false, HelpText = "Turn on verbose logging", Default = "true")]
        public string Verbose { get; set; }

        [Option('V', "scriptversion", Required = false, HelpText = "Wrapper script version", Default = "1.00")]
        public string version { get; set; }

        [Option('a', "on-measure-start", Required = false, HelpText = "Blocking command to execute before the " +
            "measurement period. Ideally should exclude Initialization.", Default = "true")]
        public string onMeasureStart { get; set; }

        [Option('b', "on-measure-stop", Required = false, HelpText = "Blocking command to execute after" +
            " the measurement period. Ideally should exclude deinitialization.", Default = "true")]
        public string onMeasureStop { get; set; }

        [Option('R', "results-directory", Required = false, HelpText = "Path to directory to" +
            " storeWrapper result fileWrapper log fileApplication raw resultsApplication log files", Default = "..\\output")]
        public string resultsDirectory { get; set; }
    }

    class Arguments
    {
        public static Options ParseArgument(string[] args, ref Options options)
        {
            if (args.Length != 0)
            {
                if (args[0] == "default")
                {
                    Utility.SetBatchScriptArguments(args, ref options);
                    options.IterationPause = 2000;
                    options.SeparationPause = 2000;
                    options.runs = 2;

                    
                    options.onMeasureStart = "True";
                    options.onMeasureStop = "False";
                    options.resultsDirectory = "..\\output";
                    options.Verbose = "True";

                    return options;
                }
            }

            else
            {
                Console.WriteLine($"\n\n{System.AppDomain.CurrentDomain.FriendlyName} default");
                Console.WriteLine("\n-------or---------\n");
            }
            args = Utility.CleanArguments(args);

            Options optionsParsed = new Options();

            Parser.Default.ParseArguments<Options>(args)
                   .WithParsed<Options>(opt =>
                   {
                       optionsParsed = opt;
                   });
            options = optionsParsed;
            
            return options;
        }
    }
}
