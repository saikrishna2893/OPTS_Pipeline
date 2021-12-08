Prerequisite:
For Creation of new project,
1) Install CommandLineParser -Version 2.8.0, ILMerge.MSBuild.Tasks -Version 1.0.0.3
 via Nuget manager.
2) Include Reference libraries: Powerpoint COM library , System.Management, System.Web.Extensions

For existing sln file,
1) Open sln file
2) Build the sln file in Release mode.
3) Get Command line arguments details by running,

   Powerpoint_AddImage.exe default
         or
   Powerpoint_AddImage.exe with below argmuents

   -i, --InputFileName               Required. Absolute path of Input filename or Relative path (In Current directory
                                    only)

  -t, --InputImageFile              Required. Absolute path of Input image filename or Relative path (In Current
                                    directory only)

  -s, --SeparationPause             (Default: 2000) Pause before and after iteration

  -p, --IterationPause              (Default: 2000) Pause between iterations

  -o, --OutputFileName              Required. Absolute path of Output filename or Relative path (In Current directory
                                    only)

  -n, --Iterations                  Required. Number of iterations

  --TargetSlideNumberList           (Default: 4,7,2,3,8,9,10,12,15,20) Target slide where image is to be pasted

  -b, --ImageAdjustingSteps         (Default: 50) Adjusting steps for Image addition to powerpoint

  -k, --TextTypingKeystrokeDelay    (Default: 50) Keystroke delay

  --StartupPause             (Default: 2000) Sleep time between application start and workload execution

  --Display                  (Default: 2) Display fullscreen setting: 1 (Full Screen) or 2 (custom window)

  --DisplayHeight            (Default: 700) Display screen height

  --DisplayWidth             (Default: 1200) Display screen width

  -r, --runs                 (Default: 1) Number to times run the application

  -v, --verbose              (Default: true) Turn on verbose logging

  -V, --scriptversion        (Default: 1.00) Wrapper script version

  -a, --on-measure-start     (Default: true) Blocking command to execute before the measurement period. Ideally should
                             exclude Initialization.

  -b, --on-measure-stop      (Default: true) Blocking command to execute after the measurement period. Ideally should
                             exclude deinitialization.

  -R, --results-directory    (Default: ..\output) Path to directory to storeWrapper result fileWrapper log
                             fileApplication raw resultsApplication log files

  --help                     Display this help screen.

  --version                  Display version information.