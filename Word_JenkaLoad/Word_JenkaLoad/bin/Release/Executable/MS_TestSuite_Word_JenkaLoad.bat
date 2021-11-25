@echo off
echo MSOffice Test Suite

SETLOCAL ENABLEDELAYEDEXPANSION
set workloadPath=.\
set inputtxtfile=..\input\Word_JenkaLoad.txt
set wprpprofilepath1=..\..\..\..\..\common\OfficeSuiteWprp.wprp
set wprpprofilepath2=..\..\common\OfficeSuiteWprp.wprp
set binary=Word_JenkaLoad
set workloadname=JenkaLoad
rem set wprprecord = 1 for recording events and generating etl file, set  wprprecord = 0 for not recording events 
set /a wprprecord=1
set wprbinary="C:\\Program Files (x86)\\Windows Kits\\10\\Windows Performance Toolkit\\wpr.exe"
SET "timestamp=%date:~10,4%%date:~4,2%%date:~7,2%-%time:~0,2%%time:~3,2%%time:~6,2%"
SET "outputdir=..\output\"
SET "resultdir=..\..\..\results\"
SET "savedir= ..\output\OPTS-%timestamp%"
SET savedir=%savedir: =%
set default=0
set /a invalidoptionlimit=3

if exist !wprpprofilepath1! (
	set wprpprofilepath=%wprpprofilepath1%
) else (
	set wprpprofilepath=%wprpprofilepath2%
)

if not exist %outputdir% mkdir %outputdir%

pushd %CD%
cd %workloadPath%

if !default!==1 (
   Set option=1
   )
if !default!==2 (
   Set option=2
)

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
   
   if !option!==1 (      
     if %wprprecord% == 1 Call !wprbinary! -start !wprpprofilepath!     
     Call !binary! default > %outputdir%!workloadname!_default.log
     if %wprprecord% == 1 Call !wprbinary! -stop %outputdir%!binary!.etl

     ECHO Cleaning Up
     mkdir %savedir%
     move %outputdir%!binary!_*.csv %savedir%
     move %outputdir%!binary!_*.json %savedir%
     move %outputdir%*!workloadname!_*.log %savedir%
     move %outputdir%!binary!*.etl %savedir%
     move %outputdir%*!workloadname!*.docx* %savedir%

   ) else (
     set CmdLine=!binary!
     echo !CmdLine!
     set /a iterationCount = 1
     
     for /f "tokens=*  usebackq delims= " %%a in (`"findstr /n ^^ %inputtxtfile% "`) do ( 
        rem check for new line
        set "var=%%a"
        set "var=!var:*:=!"     
        if not defined var  (  
          rem new line is encountered - Run the workload
          if %wprprecord% == 1 Call !wprbinary! -start %~dp0!wprpprofilepath!       
          call !CmdLine! > %outputdir%!workloadname!_Custom_!iterationCount!.log
          if %wprprecord% == 1 Call !wprbinary! -stop %outputdir%!binary!_!iterationCount!.etl            
          set /a iterationCount+=1
          echo.
          rem Initialize the CmdLine variable to the binary for next Input Config
          echo Next TestCase -----       
          set CmdLine=!binary! 
        ) else (
             rem check for comments
             set temp=!var:~0,1!     rem check the first letter of the line
             if !temp!==!key! (     
             rem do nothing - move to next line
             ) else (
          call set "CmdLine=%%CmdLine%% !var!" ) )
      )
     rem run the final set of argument
     if %wprprecord% == 1 Call !wprbinary!  -start %~dp0!wprpprofilepath!        
     call !CmdLine! > %outputdir%!workloadname!_Custom_!iterationCount!.log
     if %wprprecord% == 1 Call !wprbinary! -stop  %outputdir%!binary!_!iterationCount!.etl        
     set /a iterationCount+=1
     ECHO Cleaning Up
     mkdir %savedir%
     move %outputdir%!binary!_*.csv %savedir%
     move %outputdir%!binary!_*.json %savedir%
     move %outputdir%*!workloadname!_*.log %savedir%
     move %outputdir%!binary!*.etl %savedir%
     move %outputdir%*!workloadname!*.docx* %savedir%
     
     echo.
     echo. 
     
     
     )

if exist %resultdir% (
	if not exist %resultdir%\!binary!\ mkdir %resultdir%\!binary!\
	move %savedir% %resultdir%\!binary!
	rmdir %outputdir%
)

ECHO Done.
popd