using System;
using System.IO;
using System.Threading;
using System.Diagnostics;
using System.Collections.Generic;

using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using CommandLine;

using Helper;

namespace Workload
{
    public class Options : CommonOptions
    {
        [Option('t', "InputImageFile", Required = true, HelpText = "Absolute path of Input image filename or Relative path (In Current directory only)")]
        public string InputImageFile { get; set; }

        [Option("TargetSlideNumberList", Required = false, HelpText = "Target slide where image is to be pasted", Default = "4,7,2,3,8,9,10,12,15,20")]
        public string TargetSlideNumberList { get; set; }

        [Option('b', "ImageAdjustingSteps", Required = false, HelpText = "Adjusting steps for Image addition to powerpoint", Default = 50)]
        public int ImageAdjustingSteps { get; set; }

        [Option('k', "TextTypingKeystrokeDelay", Required = false, HelpText = "Keystroke delay", Default = 50)]
        public int TextTypingKeystrokeDelay { get; set; }

        public const bool inputFileFlag = true;
        public const bool outputFileFlag = true;
    }

    class SpecificArguments
    {
        // Command line option for AddImage
        public static Options ParseArgument(string[] args, Logger.Settings set)
        {
            Options options = new Options();

            if (args.Length != 0)
            {
                Utility.SetCaseID(args, set);

                // Invoke AddImage default
                if (args[0] == "default")
                {
                    options.InputFileName = "..\\input\\source_presentation.pptx";
                    options.InputImageFile = "..\\input\\GettyImage.jpg";
                    options.OutputFileName = "AddImageResult.pptx";
                    options.Iterations = 10;
                    options.IterationPause = 2000;
                    options.runs = 1;
                    options.TargetSlideNumberList = "4,7,2,3,8,9,10,12,15,20";
                    options.SeparationPause = 1000;
                    options.ImageAdjustingSteps = 50;
                    options.TextTypingKeystrokeDelay = 50;
                }
            }
            Arguments.ParseArgument(args, ref options);
            return options;
        }
    }


    static class AddImage
    {
        static void CreateCSV(Options opt, Logger.LogData logData, string status)
        {
            string operationname = logData.Benchmark.name;
            string startTime = logData.StartTime.time;
            string sysPath = Path.GetFullPath(opt.resultsDirectory);
            string csvfile = sysPath + "\\" + operationname + "_" + startTime + ".csv";

            string[] lines = File.ReadAllLines(csvfile);

            if (lines.Length == 0)
            {
                throw new InvalidOperationException("The file is empty");
            }
            string[] targetSlideNumberList = opt.TargetSlideNumberList.Split(',');
            //add new column to the header row
            lines[0] += ",InputFileName,TargetSlideNumber,ImageAdjustingSteps,TextTypingKeystrokeDelay";
           

            //add new column value for each row.
            for (int i = 1; i < lines.Length; i++)
            {
                lines[i] += "," + Path.GetFileName(opt.InputFileName) + "," + targetSlideNumberList[i-1] + "," + opt.ImageAdjustingSteps + "," + opt.TextTypingKeystrokeDelay;
            }
            
            //write the new content
            File.WriteAllLines(csvfile, lines);
            Logger.LogProgramEnd(status, opt);
        }

        static void ValidateIterationValues(out string[] slide, Logger.LogData logData, Options opt)
        {
            slide = opt.TargetSlideNumberList.Split(',');

            // Check the total number of iterations and columns number 
            if (opt.Iterations == -1)
            {
                opt.Iterations = slide.Length;
            }
            else if(opt.Iterations > slide.Length)
            {
                string[] slidemod = new string[opt.Iterations];
                for(int i = 0; i < opt.Iterations; i++)
                {
                    slidemod[i] = slide[(i % slide.Length)];
                }
                slide = slidemod;
                opt.TargetSlideNumberList = string.Join(",", slidemod);
            }
            else if (opt.Iterations < slide.Length)
            {
                string[] slidemod = new string[opt.Iterations];
                for (int i = 0; i < opt.Iterations; i++)
                {
                    slidemod[i] = slide[i];
                }
                slide = slidemod;
                opt.TargetSlideNumberList = string.Join(",", slidemod);
            }
        }

        static void ValidateInput(int targetSlideNumbers, int slidesCount, Logger.LogData logData, Options opt)
        {
            if(targetSlideNumbers > slidesCount + 1 || targetSlideNumbers < 1)
            {
                Console.WriteLine("Invalid slide Number. Please enter a value from 1 to slidecount+1");
                logData.Logging.Add(Logger.LogLogging(LogLevel: "Error", TimeStamp: DateTime.Now.ToString(), Detail: "Invalid slide Number. Please enter a value from 1 to slidecount+1 "));
                throw new IndexOutOfRangeException("Invalid slide Number ");
            }
        }

        static void ValidateInputImage(Options opt, Logger.LogData logData)
        {
            string InputFileName = opt.InputImageFile;

            InputFileName = Path.GetFullPath(InputFileName);
            // Check if input file exists
            if (File.Exists(InputFileName) == false)
            {
                Console.WriteLine($"Input filename {Path.GetFileName(InputFileName)} doesn't exists. Program failed");
                Logger.ExceptionDeInit("File not found Exception", $"Input filename {InputFileName} does not exists", logData, opt);
            }

            opt.InputImageFile = InputFileName;
        }

        static void PowerPointAddImage(PowerPoint.Application app, List<PowerPoint.Presentation> presentations, List<PowerPoint.Slide> ppslides, Logger.Settings set,
            Options opt, Logger.LogData logData, int repeat)
        {
            for (int rep = 1; rep <= repeat; rep++)
            {
                string operationName = "AddImage" + "_" + rep.ToString();
                set.repetition = rep;
                string[] targetSlideNumberList = null;

                ValidateIterationValues(out targetSlideNumberList, logData, opt);
                set.fileNames.Add(opt.InputFileName);
                Utility.ValidateOutputFiles(rep, opt, logData);

                try
                {
                    Utility.PowerPointInit(out app, out presentations, out ppslides, operationName, set, opt,logData);
                    set.iterationTimings.Add(operationName, new List<double>());

                    int[] targetSlideNumbers = Array.ConvertAll(targetSlideNumberList, int.Parse);

                    for (int iter = 0; iter < opt.Iterations; iter++)
                    {
                        ValidateInput(targetSlideNumbers[iter], presentations[0].Slides.Count, logData, opt);
                        Stopwatch watch = new Stopwatch();
                        Thread.Sleep(opt.SeparationPause);

                        PowerPoint.Slides slides = presentations[0].Slides;
                        int slidesCount = slides.Count;
                        PowerPoint.Slide slide = null;
                        PowerPoint.CustomLayout pictureWithCaption = slides[slidesCount - 1].CustomLayout;
                        slide = slides.AddSlide(targetSlideNumbers[iter], pictureWithCaption);

                        slide.Select();
                        PowerPoint.Shape textShape = null;
                        textShape = slide.Shapes[1];
                        slide.Shapes[1].TextFrame.TextRange.Text = $"{Path.GetFileName(opt.InputImageFile)}";

                        Thread.Sleep(opt.SeparationPause);
                        Logger.IterationEventStart(operationName, rep, iter);
                        watch.Start();

                        PowerPoint.Shape background = slide.Shapes.AddPicture(opt.InputImageFile, MsoTriState.msoFalse, MsoTriState.msoCTrue, 0, 0);
                        background.LockAspectRatio = MsoTriState.msoTrue;

                        float originalWidth = background.Width;
                        float originalHeight = background.Height;
                        float widthGap = slide.Master.Width - originalWidth;

                        for (int i = opt.ImageAdjustingSteps - 1; i >= 0; --i)
                        {
                            background.Width = slide.Master.Width - (i * widthGap / opt.ImageAdjustingSteps);
                        }

                        // The crop is applied to the original image, not the scaled one.
                        float totalCrop = (background.Height - originalHeight) * (originalWidth / background.Width);

                        for (int i = opt.ImageAdjustingSteps - 1; i >= 0; --i)
                        {
                            background.PictureFormat.CropBottom = totalCrop - (i * totalCrop / opt.ImageAdjustingSteps);
                        }

                        float centeringIncrement = -50.0f / opt.ImageAdjustingSteps;

                        for (int i = opt.ImageAdjustingSteps - 1; i >= 0; --i)
                        {
                            background.PictureFormat.Crop.PictureOffsetY += centeringIncrement;

                            watch.Stop();
                            Thread.Sleep(opt.TextTypingKeystrokeDelay);
                            watch.Start();
                        }

                        background.ZOrder(MsoZOrderCmd.msoSendToBack);

                        watch.Stop();
                        Logger.IterationEventEnd(operationName, rep, iter);

                        Thread.Sleep(opt.IterationPause);
                        set.iterationTimings[operationName].Add(watch.ElapsedMilliseconds / 1000.0);
                        Utility.IterationEnd(operationName, iter, logData, set, opt.IterationPause);
                    }

                    Thread.Sleep(opt.SeparationPause);

                    // Save the presentation
                    presentations[0].SaveAs(Utility.GetFileName(opt.OutputFileName, rep, opt.Iterations));
                    Thread.Sleep(opt.SeparationPause);

                    Utility.PowerPointDeInit(app, presentations, ppslides, set, opt.SeparationPause);
                    set.status = "success";

                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception raised. Program failed");
                    Console.WriteLine(e);
                    Console.WriteLine(e.Message);

                    Utility.PowerPointDeInit(app, presentations, ppslides, set, opt.SeparationPause);
                    set.status = "failure";

                    return;

                }
                app = null;
                presentations = null;
                ppslides = null;
            }
            Thread.Sleep(opt.SeparationPause);

            set.timingFormat = Utility.LogTimingFormat.LogGeoMeanTimingNormal;
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

            string benchMarkTestName = "Powerpoint_AddImage";
            Logger.LogData logData = Logger.Init(opt, benchMarkTestName,"Powerpoint");
            Utility.ValidateInputFiles(opt, logData);
            ValidateInputImage(opt, logData);


            // Decalre the PowerPoint objects
            PowerPoint.Application app = null;
            List<PowerPoint.Presentation> presentations = null;
            List<PowerPoint.Slide> slides = null;
            

            PowerPointAddImage(app, presentations, slides, set, opt, logData, (opt.runs));
        }
    }
}