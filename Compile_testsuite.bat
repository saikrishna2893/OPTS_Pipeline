@echo off
rem Get Current Path
set defaultPath=%CD%
set defaultPath=%defaultPath: =%

set /p version=<Version.txt
set ReleasePath=%defaultPath%\Release_%version%

set runPath=%ReleasePath%\run
set workloadsPath=%ReleasePath%\workloads
set commonPath=%workloadsPath%\common
set resultsPath=%ReleasePath%\results
echo %runPath%
echo %defaultPath%

if not exist %ReleasePath% mkdir %ReleasePath%
if not exist %runPath% mkdir %runPath%
if not exist %workloadsPath% mkdir %workloadsPath%
if not exist %resultsPath% mkdir %resultsPath%
if not exist %commonPath% mkdir %commonPath%

set Identifier=false

rem Iterate over directory
for /D %%w in (*) do (
SETLOCAL ENABLEDELAYEDEXPANSION
set workloadPath=%defaultPath%\%%~nxw

for /f "tokens=1 delims=_" %%a in ("%%~nxw") DO ( 
   set var=%%a
   
   if "!var!" == "Excel" ( set "Identifier=true")
   if "!var!" == "Word" ( set "Identifier=true")
   if "!var!" == "Outlook" ( set "Identifier=true")
   if "!var!" == "Powerpoint" ( set "Identifier=true")
   )

if "!Identifier!" == "true" (
REM Print Workload path
	echo !workloadPath!
	cd !workloadPath!
	nuget install !workloadPath!\%%~nxw\packages.config -o !workloadPath!\packages\
	echo !workloadPath!\%%~nxw.sln
	msbuild !workloadPath!\%%~nxw.sln /t:Rebuild /p:Configuration=Release /p:Platform="Any CPU"
	echo "ERROR_LEVEL" !errorlevel!
	if not !errorlevel!==0 (exit /b !errorlevel!)
	set exePath=!workloadPath!\%%~nxw\bin\Release\Executable

	if not exist !workloadsPath!\%%~nxw\bin mkdir !workloadsPath!\%%~nxw\bin
	if not exist !workloadsPath!\%%~nxw\input mkdir !workloadsPath!\%%~nxw\input
	copy !exePath! !workloadsPath!\%%~nxw\bin
	copy !workloadPath!\%%~nxw\bin\Release\input !workloadsPath!\%%~nxw\input

	cd %defaultPath%
)

copy %defaultPath%\MSOffice_Test_Automation*.bat* %runPath%
copy %defaultPath%\common\*.wprp* %commonPath%
if exist %defaultPath%\common\HelpText.txt copy %defaultPath%\common\HelpText.txt %commonPath%
if exist %defaultPath%\common\manifest.xml copy %defaultPath%\common\manifest.xml %ReleasePath%
ENDLOCAL
)

