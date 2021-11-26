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
        public int StartupPause { get; set; } = 2000;

        [Option("Display", Required = false, HelpText = "Display fullscreen setting: 1 (Full Screen) or 2 (custom window)", Default = 1)]
        public int Display { get; set; }

        [Option("DisplayHeight", Required = false, HelpText = "Display screen height ", Default = 700)]
        public double DisplayHeight { get; set; } = 700;

        [Option("DisplayWidth", Required = false, HelpText = "Display screen width ", Default = 1200)]
        public double DisplayWidth { get; set; } = 1200;

        [Option('r', "runs", Required = false, HelpText = "Number to times run the application", Default = 1)]
        public int runs { get; set; } = 1;

        [Option('v', "verbose", Required = false, HelpText = "Turn on verbose logging", Default = "true")]
        public string Verbose { get; set; } = "True";

        [Option('V', "scriptversion", Required = false, HelpText = "Wrapper script version", Default = "1.00")]
        public string scriptversion { get; set; } = "None";

        [Option('a', "on-measure-start", Required = false, HelpText = "Blocking command to execute before the " +
            "measurement period. Ideally should exclude Initialization.", Default = "True")]
        public string onMeasureStart { get; set; } = "True";

        [Option('b', "on-measure-stop", Required = false, HelpText = "Blocking command to execute after" +
            " the measurement period. Ideally should exclude deinitialization.", Default = "True")]
        public string onMeasureStop { get; set; } = "True";

        [Option('R', "results-directory", Required = false, HelpText = "Path to directory to" +
            " storeWrapper result fileWrapper log fileApplication raw resultsApplication log files", Default = "..\\output")]
        public string resultsDirectory { get; set; } = "..\\output";
                
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
                    
                    return options;
                }
            }

            else
            {
                Console.WriteLine($"\n\n{System.AppDomain.CurrentDomain.FriendlyName} default");
                Console.WriteLine("\n-------or---------\n");
            }
            string[] argsOut;
            argsOut = Utility.CleanArguments(args);
            
            Options optionsParsed = new Options();

            Parser.Default.ParseArguments<Options>(argsOut)
                   .WithParsed<Options>(opt =>
                   {
                       optionsParsed = opt;
                   });
            options = optionsParsed;
            
            return options;
        }
    }
}
