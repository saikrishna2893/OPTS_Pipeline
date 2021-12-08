Prerequisite:
For Creation of new project,
1) Install CommandLineParser -Version 2.8.0, ILMerge.MSBuild.Tasks -Version 1.0.0.3
 via Nuget manager.
2) Include Reference libraries: Excel COM library , System.Management, System.Web.Extensions

For existing sln file,
1) Open sln file
2) Build the sln file in Release mode.
3) Get Command line arguments details by running,
   Excel_Sort.exe default
         or
   Excel_Sort.exe with below argmuents

  
  -i, --InputFileName      Required. Path of Input filename or Relative path (In Current directory only)

  -p, --IterationPause     (Default: 2000) Pause between iterations

  -s, --SeparationPause    (Default: 2000) Pause before and after iteration

  --SheetNumber            (Default: 1) Sheet number eg: 1

  --SortOrder              (Default: ASC) Sort order : ASC or DES

  --Range                  (Default: A1:V400000) Range in which Sort to be perfromed eg: A1:V400000

  --ColumnList             (Default: 1,5,18,3,15,7,4,12,9,16) column numbers in the format of col1,col2,col3,col4.
                           Example: --ColumnList 1,2,3

  -o, --OutputFileName     Required. Path of Output filename or Relative path (In Current directory only)

  -n, --Iterations         (Default: -1) Number of iterations

    -r, --runs             (Default: 1) Number to times run the application
   
  --help                   Display this help screen.

  --version                Display version information.

You need run it in admin mode as by default ETW logging is enabled in batch script

MS_TestSuite_Excel_Sort could run for default and custom case(input\\Excel_Sort.txt) 