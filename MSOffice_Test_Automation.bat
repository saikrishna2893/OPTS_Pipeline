@echo off
echo MSOffice Test Suite

SETLOCAL ENABLEDELAYEDEXPANSION

set "scriptVersion=1.00"
set workloadPath=bin
rem keyword to identify comments
set key=#        
Set inputtxtfile=..\\input\\
set CmdLine=
set wprbinary="wpr.exe"
SET "timestamp=%date:~10,4%%date:~4,2%%date:~7,2%-%time:~0,2%%time:~3,2%%time:~6,2%"
SET "workloadoutputdir=..\\..\\results_%timestamp%\\"
echo %workloadUserProvidedDir%
SET "savedir=OPTS-%timestamp%"
SET savedir=%savedir: =%
set Identifier="*Excel_*","*Word_*","*Powerpoint_*","*Outlook_*"
set workloads=
rem set "csvresultfolder=..\\results\\TestSuite_CSVResults_%savedir%"
rem set "logresultfolder=..\\results\\TestSuite_LogResults_%savedir%"
set "csvresultfolder=TestSuite_CSVResults_%savedir%"
set "logresultfolder=TestSuite_LogResults_%savedir%"
rem variable to decide default/custom inputs 0-Checks with user for each workload 1-default 2-custom
set default=1
set genrateCaseID=0
set /a invalidoptionlimit=3


rem arguments to program
set "StartupPauseName=--StartupPause"
set "DisplayName=--Display"
set "DisplayHeightName=--DisplayHeight"
set "DisplayWidthName=--DisplayWidth"
set "runsName=--runs"
set "versionName=--scriptversion"
set "helpName=--help"
set "verboseName=--verbose"
set "onMeasureStartName=--on-measure-start"
set "onMeasureStopName=--on-measure-stop"
set "resultsDirectoryName=--results-directory"

set "runsNameShort=-r"
set "versionNameShort=-V"
set "helpNameShort=-h"
set "verboseNameShort=-v"
set "onMeasureStartNameShort=-a"
set "onMeasureStopNameShort=-b"
set "resultsDirectoryNameShort=-R"

set "StartupPause="
set "Display="
set "DisplayHeight="
set "DisplayWidth="
set "runs="
set "version="
set "help="
set "verbose="
set "onMeasureStart="
set "onMeasureStop="
set "resultsDirectory="

if not exist %workloadoutputdir% mkdir %workloadoutputdir%
!mkdir ..\\results\\%csvresultfolder%
!mkdir ..\\results\\%logresultfolder%

set argC=0

for %%x in (%*) do ( 
Set /A argC+=1
set "workloads=!workloads!,%%x"
)

set "workloads=!workloads:~1!"

rem Incase if user has given prefered workloads we are setting identifier
if NOT !argC!==0 (
   for /f "tokens=1,2,3,4 delims=," %%w in ("Excel_,Word_,Powerpoint_,Outlook_") DO ( 
                  
      echo !workloads!|find "%%w" > nul      
      if errorlevel 1 ( echo "found excel" > nul 2>&1  ) else ( 
         set "Identifier=!workloads!" 
         set /a genrateCaseID=1 )

      echo !workloads!|find "%%x" > nul      
      if errorlevel 1 ( echo "found word" > nul 2>&1 set /a genrateCaseID=1 ) else ( 
         set "Identifier=!workloads!" 
         set /a genrateCaseID=1 )

      echo !workloads!|find "%%y" > nul      
      if errorlevel 1 ( echo "found powerpoint" > nul 2>&1 set /a genrateCaseID=1 ) else ( set "Identifier=!workloads!"
         set /a genrateCaseID=1)

      echo !workloads!|find "%%z" > nul      
      if errorlevel 1 ( echo "found outlook" > nul 2>&1 set /a genrateCaseID=1 ) else ( set "Identifier=!workloads!"
         set /a genrateCaseID=1)
   
   )
   
   )


set "arguments=%*"
set "substring=--Display"

set "assign=None"
set "assignKeyStartUp=assignStartUpPause"
set "assignKeyDisplay=assignDisplay"
set "assignKeyDisplayHeight=assignDisplayHeight"
set "assignKeyDisplayWidth=assignDisplayWidth"
set "assignKeyRuns=assignRuns"
set "assignKeyVersion=assignVersion"
set "assignKeyHelp=assignHelp"
set "assignKeyVerbose=assignVerbose"
set "assignKeyOnMeasureStart=assignOnMeasureStart"
set "assignKeyOnMeasureStop=assignOnMeasureStop"
set "assignKeyResultsDirectory=assignResultsDirectory"


for %%A in (%arguments%) do (
   
   set currItem=%%A
   
   if !assign!==!assignKeyStartUp!   ( set "StartupPause=!currItem!" )
   if !assign!==!assignKeyDisplay!   ( set "Display=!currItem!" )
   if !assign!==!assignKeyDisplayHeight!   (      
      set "DisplayHeight=!currItem!"      
   )
   if !assign!==!assignKeyDisplayWidth!   (      
      set "DisplayWidth=!currItem!"      
   )
   if !assign!==!assignKeyRuns!   (      
      set "runs=!currItem!"      
   )
   
   if !assign!==!assignKeyHelp!   (      
      set "help=!currItem!"      
   )
   if !assign!==!assignKeyVerbose!   (      
      set "verbose=!currItem!"      
   )
   if !assign!==!assignKeyOnMeasureStart!   (      
      set "onMeasureStart=!currItem!"      
   )
   if !assign!==!assignKeyOnMeasureStop!   (      
      set "onMeasureStop=!currItem!"      
   )
   if !assign!==!assignKeyResultsDirectory!   (      
      set "resultsDirectory=!currItem!"      
   )
      
   set "assign=None"

   if !currItem!==!StartupPauseName!   (
      set "assign=assignStartUpPause"      
   )

   if !currItem!==!DisplayName!   (
      set "assign=assignDisplay" 
   )   
   if !currItem!==!DisplayHeightName!   (
      set "assign=assignDisplayHeight"
   )   
   if !currItem!==!DisplayWidthName!   (
      set "assign=assignDisplayWidth"      
   )

	if !currItem!==!runsName!   ( set "assign=assignRuns"   )   
   if !currItem!==!runsNameShort! ( set "assign=assignRuns"  )

   
   if !currItem!==!helpName!   ( set "assign=assignHelp" )
   if !currItem!==!helpNameShort!   ( set "assign=assignHelp" )

	if !currItem!==!verboseName!   ( set "assign=assignVerbose" )   
   if !currItem!==!verboseNameShort!   ( set "assign=assignVerbose" )   

   if !currItem!==!onMeasureStartName!   ( set "assign=assignOnMeasureStart" )   
   if !currItem!==!onMeasureStartNameShort!   ( set "assign=assignOnMeasureStart" )   

   if !currItem!==!onMeasureStopName!   ( set "assign=assignOnMeasureStop" )
   if !currItem!==!onMeasureStopNameShort!   ( set "assign=assignOnMeasureStop" )

	if !currItem!==!resultsDirectoryName!   ( set "assign=assignResultsDirectory" ) 
   if !currItem!==!resultsDirectoryNameShort!   ( set "assign=assignResultsDirectory" )   
    

)


if !default!==1 (
   Set option=1
   )
if !default!==2 (
   Set option=2
)

set /a casenum=100

for /f "tokens=1 delims=--" %%a in ("%Identifier%") do (
  set AFTER_UNDERSCORE=%%a  
  set Identifier=%%a
)

if !help!==True ( type ..\workloads\common\HelpText.txt )
if !help!==True ( exit /b 1 )

cd ..\\workloads\\

set workloadUserProvidedDir="..\\output"

IF not [%resultsDirectory%] == [] ( set "workloadUserProvidedDir=%resultsDirectory%" )

set "defaultArguments="
for %%a in ("!StartupPauseName!=!StartupPause! " 
"!DisplayName!=!Display! "
"!DisplayHeightName!=!DisplayHeight! "
"!DisplayWidthName!=!DisplayWidth! "
"!runsName!=!runs! "
"!versionName!=!scriptVersion! "
"!verboseName!=!verbose! "
"!onMeasureStartName!=!onMeasureStart! "
"!onMeasureStopName!=!onMeasureStop! "
"!resultsDirectoryName!=!resultsDirectory! "
) do set defaultArguments=!defaultArguments!%%~a

rem echo !defaultArguments!
rem exit /b 1

set welcome=%time%
REM echo %welcome%

rem iterate over the workloads folder
::for /D %%W in ("*Excel_*","*Word_*","*Powerpoint_*","*Outlook_*") DO (  
for /D %%W in (%Identifier%) DO (  
   echo Workload %%W 
   pushd %CD%
   rem echo %cd% 
   rem !mkdir %%W\\output
   rem cd %%W\\%%W\\%workloadPath%
   cd %%W\\%workloadPath%
   rem echo %%W\\%workloadPath%
   rem echo %cd% 
   
   set binary=%%W
   if !default!==0 (
   Set /p "option=Choose Input type 1.Default 2.Custom"
   )
   
   rem check user option
   if !option! GTR !invalidoptionlimit! ( 
      echo !option! "Invalid option, hence setting to default" 
      Set /a "option=1"      
      )
   SET "var="&for /f "delims=12" %%i in ("!option!") do set var=%%i
   if defined var ( 
      echo  !option! "Invalid option, hence setting to default"
      Set /a "option=1"    
      
   ) else ( 
      echo !option! is choosen
   )
   
   set benchmarkname=null   
   for /F "tokens=1*  delims=_" %%a in ("!binary!") do (
       set benchmarkname=%%b
   )
   
   set "STARTTIME=!time!"
   REM echo !STARTTIME!
   
   if !option!==1 ( 
     REM set "defaultArguments=!StartupPauseName!=!StartupPause! !DisplayName!=!Display! !DisplayHeightName!=!DisplayHeight! !DisplayWidthName!=!DisplayWidth! !runsName!=!runs! !versionName!=!scriptVersion! !verboseName!=!verbose! !onMeasureStartName!=!onMeasureStart! !onMeasureStopName!=!onMeasureStop! !resultsDirectoryName!=!resultsDirectory! " 
     REM echo Call !binary! default caseID-!casenum! !defaultArguments! 
     Call !wprbinary! -start %~dp0\\..\\workloads\\common\OfficeSuiteWprp.wprp
     Call !binary! default caseID-!casenum! !defaultArguments!  > !binary!_default.log
     Call !wprbinary! -stop !binary!.etl
     REM exit /b 1
     
   ) else (
      REM set "defaultArguments=!StartupPauseName!=!StartupPause! !DisplayName!=!Display! !DisplayHeightName!=!DisplayHeight! !DisplayWidthName!=!DisplayWidth!" 
      
     set CmdLine=!binary!
     echo !CmdLine!
     set "CmdLine=!CmdLine! caseID-!casenum! !defaultArguments!"
     set /a iterationCount = 1
    
     
     for /f "tokens=*  usebackq delims= " %%a in (`"findstr /n ^^ %inputtxtfile%%%W.txt "`) do ( 
        REM check for new line
        set "var=%%a"
        set "var=!var:*:=!"     
        if not defined var  (  
          REM new line is encountered - Run the workload      
          Call !wprbinary! -start %~dp0\\..\\workloads\\common\OfficeSuiteWprp.wprp          
          call !CmdLine!   > !binary!_Custom_!iterationCount!.log
          Call !wprbinary! -stop !binary!_!iterationCount!.etl
          set /a iterationCount+=1
          echo.
          REM Initialize the CmdLine variable to the binary for next Input Config
          echo Next TestCase -----
          if %genrateCaseID%==1 set /a casenum=casenum+100       
          set CmdLine=!binary!
          set "CmdLine=!CmdLine! caseID-!casenum!  !defaultArguments!" 
        ) else (
             REM check for comments
             set temp=!var:~0,1!     rem check the first letter of the line
             if !temp!==!key! (     
             REM do nothing - move to next line
             ) else (
          call set "CmdLine=%%CmdLine%% !var!" ) )
      )
     REM run the final set of argument
     Call !wprbinary! -start %~dp0\\..\\workloads\\common\OfficeSuiteWprp.wprp
     call !CmdLine!   > !binary!_Custom_!iterationCount!.log
     Call !wprbinary! -stop !binary!_!iterationCount!.etl

     set /a iterationCount+=1
     echo.
     echo. )
     
	 set "ENDTIME=!time!"
	 REM echo End Time---
	 REM echo !ENDTIME!
	 
	 rem Change formatting for the start and end times
    for /F "tokens=1-4 delims=:.," %%a in ("!STARTTIME!") do (
       set /A "start=(((%%a*60)+1%%b %% 100)*60+1%%c %% 100)*100+1%%d %% 100"
    )
    for /F "tokens=1-4 delims=:.," %%a in ("!ENDTIME!") do ( 
       IF "!ENDTIME!" GTR "!STARTTIME!" set /A "end=(((%%a*60)+1%%b %% 100)*60+1%%c %% 100)*100+1%%d %% 100" 
       IF "!ENDTIME!" LSS "!STARTTIME!" set /A "end=((((%%a+24)*60)+1%%b %% 100)*60+1%%c %% 100)*100+1%%d %% 100" 
    )
    rem Calculate the elapsed time by subtracting values
    set /A "elapsed=!end!%%-!start!"
	
    rem Format the results for output
    set /A "hh=!elapsed!/(60*60*100)"
	set /A "rest=!elapsed!%%(60*60*100)"
	set /A "mm=!rest!/(60*100)"
	set /A "sec=!rest!%%(60*100)"
	set	/A "ss=!sec!/100"
	set	/A "cc=!sec!%%(200)"
    
    set Elapsed=!hh!h:!mm!m:!ss!s.!cc!ms
	
	REM echo Start    : !STARTTIME!
    REM echo Finish   : !ENDTIME!
    REM echo          ---------------
    echo %%W - Elapsed time : !Elapsed! >> %~dp0..\\results\\TimingLog-!savedir!.txt
	 
     rem Copying CSV to target results folder in root
     rem first it fetches all csv's in a directory and then checks if there is a number after "-" if its so it will copy to target folder
     rem we are using "-" because usually the format is {Workload}_date-time.csv
     set timestamp=null
     rem echo "csv files"
     echo %cd%
     for %%Q in (!workloadUserProvidedDir!\\*.csv) do (        
          for /F "tokens=1*  delims=-" %%a in ("%%~nQ") do (
             SET "csv_var="&for /f "delims=0123456789" %%i in ("%%b") do set csv_var=%%i
             if NOT "%%b"=="" (
                if NOT defined csv_var (
                     
				        copy %%Q  %~dp0..\\results\\%csvresultfolder%
                    type %%Q  >> %~dp0..\\results\\%csvresultfolder%_Copy.csv
                )
             )
          
         )
     )
	 
	 for %%Q in (*.log) do (        
          
        type %%Q  >> %~dp0..\\results\\%logresultfolder%.log               
             
     )
	 
	  copy !binary!_*.log %~dp0..\\results\\%logresultfolder%\\
     timeout /t 2 >nul
     
     rem Cleaning up files
     ECHO Cleaning Up
     mkdir %savedir%     
     timeout /t 2 >nul
     
     move !workloadUserProvidedDir!\\!binary!_*.json %savedir%
     move !binary!_*.log %savedir%
     move !binary!*.etl %savedir%
     move !workloadUserProvidedDir!\\!binary!*.csv %savedir%
     move !workloadUserProvidedDir!\\*!benchmarkname!*.x* %savedir%  > nul 2>&1
     move !workloadUserProvidedDir!\\*!benchmarkname!*.pdf* %savedir% > nul 2>&1
     move !workloadUserProvidedDir!\\*!benchmarkname!*.csv* %savedir% > nul 2>&1
     move !workloadUserProvidedDir!\\*!benchmarkname!*.doc* %savedir% > nul 2>&1
     move !workloadUserProvidedDir!\\*!benchmarkname!*.pp* %savedir% > nul 2>&1
	 move !workloadUserProvidedDir!\\*!benchmarkname!*.msg* %savedir% > nul 2>&1
	 move !workloadUserProvidedDir!\\*!benchmarkname!*.jpg* %savedir% > nul 2>&1
	 move !workloadUserProvidedDir!\\*!benchmarkname!*.vcf* %savedir% > nul 2>&1
	 move !workloadUserProvidedDir!\\*!benchmarkname!*.pot* %savedir% > nul 2>&1
	 move !workloadUserProvidedDir!\\*!benchmarkname!*.thm* %savedir% > nul 2>&1
	 move !workloadUserProvidedDir!\\*!benchmarkname!*.mp* %savedir% > nul 2>&1

     !mkdir %~dp0\\..\\results\\!binary!     
     move %savedir%  %~dp0\\..\\results\\!binary!
	 rmdir "!workloadUserProvidedDir!" > nul 2>&1
     rem mkdir %savedir% 
     
	 timeout /t 2 >nul
   popd

   rem new case id for next workload
   if %genrateCaseID%==1 set /a casenum=casenum+100  

  )
  
  set endTime=%TIME%

    rem ******************  END MAIN CODE SECTION

    rem Change formatting for the start and end times
    for /F "tokens=1-4 delims=:.," %%a in ("%welcome%") do (
       set /A "start=(((%%a*60)+1%%b %% 100)*60+1%%c %% 100)*100+1%%d %% 100"
    )

    for /F "tokens=1-4 delims=:.," %%a in ("%endTime%") do ( 
       IF %endTime% GTR %welcome% set /A "end=(((%%a*60)+1%%b %% 100)*60+1%%c %% 100)*100+1%%d %% 100" 
       IF %endTime% LSS %welcome% set /A "end=((((%%a+24)*60)+1%%b %% 100)*60+1%%c %% 100)*100+1%%d %% 100" 
    )

    rem Calculate the elapsed time by subtracting values
    set /A elapsed=end-start

    rem Format the results for output
    set /A hh=elapsed/(60*60*100), rest=elapsed%%(60*60*100), mm=rest/(60*100), rest%%=60*100, ss=rest/100, cc=rest%%100
    REM if %hh% lss 10 set hh=0%hh%
    REM if %mm% lss 10 set mm=0%mm%
    REM if %ss% lss 10 set ss=0%ss%
    REM if %cc% lss 10 set cc=0%cc%

    set DURATION=%hh%h:%mm%m:%ss%s.%cc%ms

    echo Start    : !welcome!
    echo Finish   : %endTime%
    echo          ---------------
    echo Total TestSuite Duration : %DURATION% >> %~dp0..\\results\\TimingLog-%savedir%.txt

timeout /t 2
set "tempcsvfilename=%~dp0\\..\\results\\%csvresultfolder%_Copy.csv"
rem in case if user feeds a list  

if %genrateCaseID%==1 (
   set /a count=1 
   (for /f "usebackq tokens=1-14 delims=," %%a in (!tempcsvfilename!) do ( 
      rem echo %%a     
      if %%a==Case_ID if !count!==1 echo %%a,%%b,%%c,%%d,%%e,%%f,%%g,%%h,%%i,%%j,%%k,%%l,%%m,%%n
      if NOT %%a==Case_ID echo %%a,%%b,%%c,%%d,%%e,%%f,%%g,%%h,%%i,%%j,%%k,%%l,%%m,%%n
      set /a count=count+1
      )) > "%~dp0\\..\\results\\%csvresultfolder%.csv"
) else (   
   rename !tempcsvfilename! %csvresultfolder%.csv > nul 2>&1
)
timeout /t 4

del !tempcsvfilename! > nul 2>&1
  

 
 
ENDLOCAL

