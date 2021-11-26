### Microsoft Office TestSuite Suite 

## Pre-requistie
 
1. Microsoft Office Installation

## Procedure

**For recording Windows Performance**

1. Install Windows Performance kit and Windows Performance Recorder

**Compilation Script**

1. Set two paths via in command prompt.
	1. For VS Tools  
	2. For nuget.exe 

	Example
	1. set PATH=%PATH%;C:\Program Files (x86)\Microsoft Visual Studio\2019\Community\Common7\Tools;
	2. set PATH=%PATH%;C:\Users\amduser\...\msoffice_test_suite;  (if nuget.exe is placed in msoffice_test_suite)
	
2. Navigate VSTools path in command prompt and execute VsMSBuildCmd.bat from 
3. Run Compile_testsuite.bat from root folder

**Boss Script**
1. Open Admin cmd prompt , Go to root folder

2. Run MSOffice_Test_Automation.bat

	- If you want to run a specific workload use MSOffice_Test_Automation.bat Excel_FormatTable - common columns will be logged in root folder csv
	- if you want a specific set of workloads use MSOffice_Test_Automation.bat Excel_FormatTable,Excel_AppendTable - common columns will be logged in root folder csv
	- if you want to run all workloads use MSOffice_Test_Automation.bat 

	Notes 

	- if you want run default case for all workloads you can set "set default=1" in MSOffice_Test_Automation.bat
	- if you want run custom case for all workloads you can set "set default=2" in MSOffice_Test_Automation.bat


**Release #2 - Features**

- Common files are created for all workloads - 
	CSVlogger - Logs csv information and created csv file
	EtwLogger - Etw event functions
	Logger.cs - Functions required for time calculation and json file creation
	Utility.cs - Validated input,outputfiles, Initalizes and deinitalizes application 
	Helper.cs - Common Commandline parameters are added here
- User can chose to set full screen or specific window size for all workloads "Display", "DisplayHeight", "DisplayWidth"
- "StartupPause" thread sleep is added after each type of workload(Excel,Word,Powerpoint,Outlook) application opening 
- Version for all workloads are initalized as 1.00 and it is logged in all results.

**Release #3 - Features**

- Compilation Script 
	- After running compilation script we will get a release directory(with version suffix) with which user can independently test all workloads without .dll files and project files.

- Added new Top Level CommandLine arguments
	- Following arguments were added --> --runs(-r), --scriptversion(-V), --help(-h), --verbose(-v), --on-measure-start(-a), --on-measure-stop(-b), --results-directory(-R)

- Added manifest file 

- Removed other branches

