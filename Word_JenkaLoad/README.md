### Office Suite - Jenka_968 source workload

* Word Workload
* Replicates Jenka's Load workload Opens document and adjusts the zooming level of document and for the first two input files alone a drop run will happen
        
**Prerequisite**

For Creation of new project,

- Install CommandLineParser -Version 2.8.0, ILMerge.MSBuild.Tasks -Version 1.0.0.3
 via Nuget manager.
- Include Reference libraries: Word COM library , System.Management, System.Web.Extensions

For existing sln file,

- Open sln file
- Build the sln file in Release mode.
- Get Command line arguments details by running,
   Word_JenkaLoad.exe default
         or
   Word_JenkaLoad.exe with below argmuents

  -i, --InputFileName     Required. Path of Input filename or Relative path (In Current directory only) can specify
                          multiple input files separated by commaseg: inputfile1,inputfile2

  -p, --IterationPause    (Default: 2000) Pause between iterations

  --SeparationPause       (Default: 2000) Pause before and after iteration

  --LoadIterationList     (Default: 1,1,1,1,1,5) Times on which each document should be loaded

  -o, --OutputFileName    Required. Path of Output Filename

  -z, --ZoomPercentage    (Default: 100) Zoom percentage

  -n, --Iterations        (Default: -1) Number of iterations (in this case it is number of test cases to be considered)

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