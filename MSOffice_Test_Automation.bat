@echo off
echo MSOffice Test Suite

SETLOCAL ENABLEDELAYEDEXPANSION

set workloadPath=bin
rem keyword to identify comments
set key=#        
Set inputtxtfile=..\\input\\
set CmdLine=
set wprbinary="wpr.exe"
SET "timestamp=%date:~10,4%%date:~4,2%%date:~7,2%-%time:~0,2%%time:~3,2%%time:~6,2%"
SET "workloadoutputdir=..\\..\\results\\"
SET "savedir=OPTS-%timestamp%"
SET savedir=%savedir: =%
set Identifier="*Excel_*","*Word_*","*Powerpoint_*","*Outlook_*"
set workloads=
set "csvresultfolder=..\\results\\TestSuite_CSVResults_%savedir%"
set "logresultfolder=..\\results\\TestSuite_LogResults_%savedir%"
rem variable to decide default/custom inputs 0-Checks with user for each workload 1-default 2-custom
set default=0
set genrateCaseID=0
set /a invalidoptionlimit=3

rem arguments to program
set "StartupPauseName=--StartupPause"
set "DisplayName=--Display"
set "DisplayHeightName=--DisplayHeight"
set "DisplayWidthName=--DisplayWidth"
set "StartupPause="
set "Display="
set "DisplayHeight="
set "DisplayWidth="

echo !workloads!

!mkdir %csvresultfolder%
!mkdir %logresultfolder%

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
      if errorlevel 1 ( echo "found excel" > nul 2>&1 ) else ( set "Identifier=!workloads!")

      echo !workloads!|find "%%x" > nul      
      if errorlevel 1 ( echo "found word" > nul 2>&1 ) else ( set "Identifier=!workloads!")

      echo !workloads!|find "%%y" > nul      
      if errorlevel 1 ( echo "found powerpoint" > nul 2>&1 ) else ( set "Identifier=!workloads!")

      echo !workloads!|find "%%z" > nul      
      if errorlevel 1 ( echo "found outlook" > nul 2>&1 ) else ( set "Identifier=!workloads!")
   
   )
   set /a genrateCaseID=1
   )

set "arguments=%*"
set "substring=--Display"

set "assign=None"
set "assignKeyStartUp=assignStartUpPause"
set "assignKeyDisplay=assignDisplay"
set "assignKeyDisplayHeight=assignDisplayHeight"
set "assignKeyDisplayWidth=assignDisplayWidth"

for %%A in (%arguments%) do (
   
   set currItem=%%A
   
   if !assign!==!assignKeyStartUp!   (      
      set "StartupPause=!currItem!"      
   )
   if !assign!==!assignKeyDisplay!   (      
      set "Display=!currItem!"      
   )
   if !assign!==!assignKeyDisplayHeight!   (      
      set "DisplayHeight=!currItem!"      
   )
   if !assign!==!assignKeyDisplayWidth!   (      
      set "DisplayWidth=!currItem!"      
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


cd ..\\workloads\\

rem echo %cd%
rem exit /b 1
rem iterate over the workloads folder
::for /D %%W in ("*Excel_*","*Word_*","*Powerpoint_*","*Outlook_*") DO (  
for /D %%W in (%Identifier%) DO (  
   echo Workload %%W 
   pushd %CD%
   rem echo %cd% 
   
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
   
   if !option!==1 ( 
     set "defaultArguments=!StartupPauseName!=!StartupPause! !DisplayName!=!Display! !DisplayHeightName!=!DisplayHeight! !DisplayWidthName!=!DisplayWidth!" 
     
     Call !wprbinary! -start %~dp0\\..\\workloads\\common\OfficeSuiteWprp.wprp
     Call !binary! default caseID-!casenum! !defaultArguments!  > !binary!_default.log
     Call !wprbinary! -stop !binary!.etl
     
     
   ) else (
      set "defaultArguments=!StartupPauseName!=!StartupPause! !DisplayName!=!Display! !DisplayHeightName!=!DisplayHeight! !DisplayWidthName!=!DisplayWidth!" 
      
     set CmdLine=!binary!
     echo !CmdLine!
     set "CmdLine=!CmdLine! caseID-!casenum! !defaultArguments!"
     set /a iterationCount = 1
    
     
     for /f "tokens=*  usebackq delims= " %%a in (`"findstr /n ^^ %inputtxtfile%%%W.txt "`) do ( 
        rem check for new line
        set "var=%%a"
        set "var=!var:*:=!"     
        if not defined var  (  
          rem new line is encountered - Run the workload      
          Call !wprbinary! -start %~dp0\\..\\workloads\\common\OfficeSuiteWprp.wprp          
          call !CmdLine!   > !binary!_Custom_!iterationCount!.log
          Call !wprbinary! -stop !binary!_!iterationCount!.etl
          set /a iterationCount+=1
          echo.
          rem Initialize the CmdLine variable to the binary for next Input Config
          echo Next TestCase -----
          if %genrateCaseID%==1 set /a casenum=casenum+100       
          set CmdLine=!binary!
          set "CmdLine=!CmdLine! caseID-!casenum!  !defaultArguments!" 
        ) else (
             rem check for comments
             set temp=!var:~0,1!     rem check the first letter of the line
             if !temp!==!key! (     
             rem do nothing - move to next line
             ) else (
          call set "CmdLine=%%CmdLine%% !var!" ) )
      )
     rem run the final set of argument
     Call !wprbinary! -start %~dp0\\..\\workloads\\common\OfficeSuiteWprp.wprp
     call !CmdLine!   > !binary!_Custom_!iterationCount!.log
     Call !wprbinary! -stop !binary!_!iterationCount!.etl

     set /a iterationCount+=1
     echo.
     echo. )
     
     rem Copying CSV to target results folder in root
     rem first it fetches all csv's in a directory and then checks if there is a number after "-" if its so it will copy to target folder
     rem we are using "-" because usually the format is {Workload}_date-time.csv
     set timestamp=null
     rem echo "csv files"
     echo %cd%
     for %%Q in (..\\output\\*.csv) do (        
          for /F "tokens=1*  delims=-" %%a in ("%%~nQ") do (
             SET "csv_var="&for /f "delims=0123456789" %%i in ("%%b") do set csv_var=%%i
             if NOT "%%b"=="" (
                if NOT defined csv_var (
                     
				        copy %%Q  %~dp0%csvresultfolder%
                    type %%Q  >> %~dp0%csvresultfolder%_Copy.csv
                )
             )
          
         )
     )
	 
	 for %%Q in (*.log) do (        
          
        type %%Q  >> %~dp0\\%logresultfolder%.log               
             
     )
	 
	  copy !binary!_*.log %~dp0\\%logresultfolder%\\
     timeout /t 2 >nul
     
     rem Cleaning up files
     ECHO Cleaning Up
     mkdir %savedir%     
     timeout /t 2 >nul
     
     move ..\\output\\!binary!_*.json %savedir%
     move !binary!_*.log %savedir%
     move !binary!*.etl %savedir%
     move ..\\output\\!binary!*.csv %savedir%
     move ..\\output\\*!benchmarkname!*.x* %savedir%  > nul 2>&1
     move ..\\output\\*!benchmarkname!*.pdf* %savedir% > nul 2>&1
     move ..\\output\\*!benchmarkname!*.csv* %savedir% > nul 2>&1
     move ..\\output\\*!benchmarkname!*.doc* %savedir% > nul 2>&1
     move ..\\output\\*!benchmarkname!*.pp* %savedir% > nul 2>&1

     !mkdir %~dp0\\..\\results\\!binary!     
     move %savedir%  %~dp0\\..\\results\\!binary!
     rem mkdir %savedir% 
     

     
	 timeout /t 2 >nul
   popd

   rem new case id for next workload
   if %genrateCaseID%==1 set /a casenum=casenum+100  

  )

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
   rename !tempcsvfilename! "%~dp0\\..\\results\\%csvresultfolder%.csv" > nul 2>&1
)
timeout /t 4
rem del !tempcsvfilename! 2>&1
  

 
 
ENDLOCAL

