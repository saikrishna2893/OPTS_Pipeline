### Office Suite - Insert Envelope

* Parallel workload.
* Inserts an Envelope with from from and to address to send physical mails
        
**Prerequisite**

For Creation of new project,

- Install CommandLineParser -Version 2.8.0, ILMerge.MSBuild.Tasks -Version 1.0.0.3
 via Nuget manager.
- Include Reference libraries: Word COM library , System.Management, System.Web.Extensions

For existing sln file,

- Open sln file
- Build the sln file in Release mode.
- Get Command line arguments details by running,
   Word_InsertEnvelope.exe default
         or
   Word_InsertEnvelope.exe with below argmuents

  -i, --InputFileName      Required. Path of Input filename or Relative path (In Current directory only)

  -p, --IterationPause     (Default: 1000) Pause between iterations

  -s, --SeparationPause    (Default: 2000) Pause between iterations

  -o, --OutputFileName     Required. Path of Output filename or Relative path (In Current directory only)

  -n, --Iterations         Required. Number of iterations

  -b, --ToAddress          (Default: Sujith L,49th Cross,Bangalore,Karnataka.Krish S,Greenways,Mumbai,Maharastra)
                           Address to send the mail (Specify in a single line with commas)

  -a, --FromAddress        (Default: Ajay K, 1st Street, Chennai, Tamil Nadu.John L, 2nd St, Hyderabad, Telangana)
                           Return/From Address (Specify in a single line with commas and while entering multiple
                           addresses separate with .)

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