@ECHO off 

:: Get arguments are passed by jenkins
SET "MSN=%~1"
SET "COMMIT=%~2"
SET "REVISION=%~3"
SET "TS_TO_TR=%~4"
@ECHO TS TO TR %TS_TO_TR%

:: Moudle
SET PATH=C:\Python27\Scripts;%PATH%
SET PATH=C:\Program Files\TortoiseSVN\bin;%PATH%
SET PATH=C:\Program Files\7-Zip;%PATH%
SET ZIP_FILENAME_REPORT=RH850_X2x_%MSN%_TUT_Result.zip
:: REPO trunk
:: SET ROOT_REPO=D:\Workspace\s979 sang D:\00_Repo\00_X2x_Trunk
SET TEST_WORKSPACE=D:\02_Jobs\UT\01_TUT
SET ROOT_REPO=D:\00_Repo\01_X2x_U2A8_FS

SET /A INDEX=0
SET DEVICE=E2x

:: Delete Running file
IF EXIST %MSN%_running.txt (
  DEL /S /Q %MSN%_running.txt
  ECHO %MSN%_running.txt is removed
)
IF EXIST svn_updating.txt (
  DEL /S /Q svn_updating.txt
  ECHO svn_updating.txt is removed
)

type nul > %MSN%_running.txt

:: MAPPING U DRIVER
IF EXIST U: SUBST U: /D
IF EXIST %ROOT_REPO% (
   SUBST U: %ROOT_REPO%
   ECHO Current virtual drive map is ....
   SUBST
) ELSE (
    ECHO %ROOT_REPO% is not existed.
    Exit /b
)
:CHECK
IF EXIST svn_updating.txt (
    ECHO Other module is updating Svn, please wait a little bit.
    goto WAIT
)

type nul > svn_updating.txt

svn cleanup --remove-unversioned "%ROOT_REPO%\external\X2x\common\generic\generator"
svn cleanup --remove-unversioned "%ROOT_REPO%\internal\Module\generic\06_CD\01_WorkProduct\generator_cs"
svn cleanup --remove-unversioned "%ROOT_REPO%\internal\Module\%MSN%\07_UT\01_WorkProduct_T"
svn cleanup --remove-unversioned "%ROOT_REPO%\internal\Module\%MSN%\06_CD\01_WorkProduct\generator_cs"
svn cleanup --remove-unversioned "%TEST_WORKSPACE%"

svn revert -R "%ROOT_REPO%\external\X2x\common\generic\generator"
svn revert -R "%ROOT_REPO%\internal\Module\generic\06_CD\01_WorkProduct\generator_cs"
svn revert -R "%ROOT_REPO%\internal\Module\%MSN%\07_UT\01_WorkProduct_T"
svn revert -R "%ROOT_REPO%\internal\Module\%MSN%\06_CD\01_WorkProduct\generator_cs"
svn revert -R "%TEST_WORKSPACE%"

svn update "%ROOT_REPO%\internal\Module\%MSN%\07_UT\01_WorkProduct_T"
svn update "%TEST_WORKSPACE%"

IF "%REVISION%"=="New" (
  svn update "%ROOT_REPO%\external"
  svn update "%ROOT_REPO%\internal\Module\generic\06_CD\01_WorkProduct\generator_cs"
  svn update "%ROOT_REPO%\internal\Module\%MSN%\06_CD\01_WorkProduct\generator_cs"
) ELSE (
  svn update -r %REVISION% %ROOT_REPO%\external
  svn update -r %REVISION% %ROOT_REPO%\internal\Module\generic\06_CD\01_WorkProduct\generator_cs
  svn update -r %REVISION% %ROOT_REPO%\internal\Module\%MSN%\06_CD\01_WorkProduct\generator_cs
)

DEL /S /Q svn_updating.txt

:: Copy all dlls from D:/test/external/X2x/common/generic/generator/dlls/E2x or D:/test/external/X2x/common/generic/generator/dlls/U2x to D:/test/external/X2x/common/generic/generator
XCOPY /y "U:\external\X2x\common\generic\generator\dlls\E2x\*.dll" "U:\external\X2x\common\generic\generator\"
XCOPY /y "U:\external\X2x\common\generic\generator\dlls\U2x\*.dll" "U:\external\X2x\common\generic\generator\"

:: Create folder workspace and copy all tool to workspace folder
mkdir U:\internal\Module\%MSN%\07_UT\01_WorkProduct_T\workspace
SET MSN_TEST_WORKSPACE=U:\internal\Module\%MSN%\07_UT\01_WorkProduct_T\workspace
XCOPY /y "%TEST_WORKSPACE%\*" "%MSN_TEST_WORKSPACE%\" /E

SET RESULT=

:LOOP
@ECHO --- Run Unitest for device: %DEVICE% ---
SET "EX_FILE=%MSN%%DEVICE%"

:: Generic module
IF "%MSN%"=="Generic" (
	SET DEVICE=%MSN%
	SET EX_FILE=%MSN%
	SET EX_FILE=%MSN%
	SET /A INDEX=2
)

::@RD /s /q U:/internal/Module/%MSN%/07_UT/01_WorkProduct_T/app/TUT_%EX_FILE%Commonize/TUT_%EX_FILE%Commonize/bin/

:: Build project Test
START "" /w /b C:/"Program Files (x86)"//"Microsoft Visual Studio"//2019/Enterprise/MSBuild/Current/Bin/MSBuild.exe "U:/internal/Module/%MSN%/07_UT/01_WorkProduct_T/app/TUT_%EX_FILE%Commonize/TUT_%EX_FILE%Commonize.sln" /p:PreBuildEvent=;PostBuildEvent=

:: IF build failed => Exit job
IF errorlevel 1 (
 SET RESULT="BUILD FAILED"
 SET COMMIT=No
 GOTO EXIT
) 

:: Run all test cases
START "" /w /b C:/"Program Files (x86)"//"Microsoft Visual Studio"/2019/Enterprise/Common7/IDE/CommonExtensions/Microsoft/TestWindow/vstest.console.exe "U:/internal/Module/%MSN%/07_UT/01_WorkProduct_T/app/TUT_%EX_FILE%Commonize/TUT_%EX_FILE%Commonize/bin/Debug/TUT_%EX_FILE%Commonize.dll" /Enablecodecoverage /Settings:"%MSN_TEST_WORKSPACE%/output.runsettings" /Logger:trx;LogFileName=Test_Result_%DEVICE%.trx

IF errorlevel 1 (
 SET RESULT="Run TCs Failed"
 SET COMMIT=No
) 

::Get file report
for /f %%f in ('Where /r %MSN_TEST_WORKSPACE%\TestResults Coverage_result.coverage') do (
    SET COVERAGE_FILE=%%f
)

IF NOT EXIST %MSN_TEST_WORKSPACE%\Output md %MSN_TEST_WORKSPACE%\Output
md %MSN_TEST_WORKSPACE%\Output\%DEVICE%

IF "%COVERAGE_FILE%"=="" GOTO EXIT

:: Convert file .coverage report to xml format file
::START "" /w /b "C:/Program Files (x86)/Microsoft Visual Studio/2017/Enterprise/Team Tools/Dynamic Code Coverage Tools/CodeCoverage.exe" analyze /output:"%MSN_TEST_WORKSPACE%\Output\%DEVICE%\%DEVICE%_DynamicCodeCoverage.coveragexml" %COVERAGE_FILE%

:: Build project Convert coverage file to coveragexml file -> to exe file
START "" /w /b C:/"Program Files (x86)"//"Microsoft Visual Studio"//2019/Enterprise/MSBuild/Current/Bin/MSBuild.exe "%MSN_TEST_WORKSPACE%/Tool/CovertCoverage/ConvertCoverage/ConvertCoverage.sln"

XCOPY /y /F "%COVERAGE_FILE%" "%MSN_TEST_WORKSPACE%/Output/%DEVICE%/"
XCOPY /y /F "%MSN_TEST_WORKSPACE%/TestResults/Test_Result_%DEVICE%.trx" "%MSN_TEST_WORKSPACE%/Output/%DEVICE%/"
DEL /S /Q "%MSN_TEST_WORKSPACE%/TestResults/

:: Copy "C:\Program Files (x86)\Microsoft Visual Studio\2017\Enterprise\Common7\IDE\PrivateAssemblies\Microsoft.VisualStudio.Coverage.Symbols.dll" to place of .coverage file
XCOPY /Y /F "C:/Program Files (x86)/Microsoft Visual Studio/2019/Enterprise/Common7/IDE/PrivateAssemblies/Microsoft.VisualStudio.Coverage.Symbols.dll" "%MSN_TEST_WORKSPACE%/Output/%DEVICE%/"

:: Covert .converage to .coveragexml
"%MSN_TEST_WORKSPACE%/Tool/CovertCoverage/ConvertCoverage/ConvertCoverage/bin/Debug/ConvertCoverage.exe" "%MSN_TEST_WORKSPACE%/Output/%DEVICE%/Coverage_result.coverage" "%MSN_TEST_WORKSPACE%\Output\%DEVICE%\%DEVICE%_DynamicCodeCoverage.coveragexml"


set /A INDEX=%INDEX%+1

IF %INDEX% == 1 (
 SET DEVICE=U2x
 GOTO LOOP
)
:: Call tool python to gen report
python %MSN_TEST_WORKSPACE%\Tool\GetCoverageReport.py %MSN% %COMMIT%

:: Zip all result
set mypath=%cd%

:: Delete files unneed
IF "%MSN%"=="Generic" (
	CD /d U:\internal\Module\%MSN%\07_UT\01_WorkProduct_T\workspace\Output\%MSN%
    DEL /S /Q %MSN%_DynamicCodeCoverage.coveragexml Test_Result_%MSN%.trx Microsoft.VisualStudio.Coverage.Symbols.dll
) ELSE (
	FOR %%i in (E2x U2x) DO (
		CD /d U:\internal\Module\%MSN%\07_UT\01_WorkProduct_T\workspace\Output\%%i
		DEL /S /Q %%i_DynamicCodeCoverage.coveragexml Test_Result_%%i.trx Microsoft.VisualStudio.Coverage.Symbols.dll
	)
)
CD /d %mypath%

:: Delete Zip file if  they are existed 
IF EXIST %ZIP_FILENAME_REPORT% (
    DEL /S /Q %ZIP_FILENAME_REPORT%
)

:: Copy TS to TR and fill PASSED
:: This function is beta => Please rem if it have problem
IF "%TS_TO_TR%"=="true" (
	IF "%RESULT%"=="" (
		START "" /w /b "%MSN_TEST_WORKSPACE%/Tool/CopyTStoTR/UpdateTestReport/bin/Debug/UpdateTestReport.exe" %MSN%
		mkdir "%MSN_TEST_WORKSPACE%\Output\test_report"
		XCOPY "U:\internal\Module\%MSN%\07_UT\01_WorkProduct_T\result\U2A8\Beta\test_report\*" "%MSN_TEST_WORKSPACE%\Output\test_report"
		DEL /S /Q %MSN_TEST_WORKSPACE%\Output\test_report\RH850_X2x_%MSN%_TUT_CR_U2A8_Beta.xlsx
	)
)

:: Zip all files in Output folder
7z a %ZIP_FILENAME_REPORT% %MSN_TEST_WORKSPACE%\Output\

:: If commit is selected then Zip coverage file, copy .coverage and log file to test_log folder. After committing them.  
IF "%COMMIT%" == "Yes" (
	IF "%MSN%"=="Generic" (
		7z a %MSN_TEST_WORKSPACE%\Output\RH850_X2x_%MSN%_TUT_CoverageResult_U2A8_Beta.zip %MSN_TEST_WORKSPACE%\Output\%MSN%\RH850_X2x_%MSN%_TUT_CoverageResult_U2A8_Beta.coverage
	) ELSE (
		7z a %MSN_TEST_WORKSPACE%\Output\RH850_U2x_%MSN%_TUT_U2A8_Beta.zip %MSN_TEST_WORKSPACE%\Output\U2x\RH850_U2x_*.coverage
		7z a %MSN_TEST_WORKSPACE%\Output\RH850_E2x_%MSN%_TUT_U2A8_Beta.zip %MSN_TEST_WORKSPACE%\Output\E2x\RH850_E2x_*.coverage
	)

    CD /d U:\internal\Module\%MSN%\07_UT\01_WorkProduct_T\result\U2A8\Beta\test_report\
    DEL /Q /S RH850_X2x_%MSN%_TUT_CR_U2A8_Beta.xlsx
    XCOPY /Y /F %MSN_TEST_WORKSPACE%\Output\RH850_X2x_%MSN%_TUT_CR_U2A8_Beta.xlsx .\
	
    ::Commit Coverage Report file
    svn commit -m "Commit via Jenkins - Coverage_result" RH850_X2x_%MSN%_TUT_CR_U2A8_Beta.xlsx

    CD ../test_log
    DEL /Q /S RH850_*.zip
    DEL /Q /S RH850_X2x*TestLog.log

    :: Rename and copy zip to log folder
    python %MSN_TEST_WORKSPACE%\Tool\rename.py %MSN%
    XCOPY /Y %MSN_TEST_WORKSPACE%\Output\RH850_*_TUT*_U2A8_Beta.zip .\

    :: Delete old coveragexml and add zip file
	IF EXIST RH850*.coveragexml (
		svn delete RH850*.coveragexml
        svn add RH850_*_TUT_U2A8_Beta.zip
	)
    
    FOR %%i in (E2x U2x) DO (
        XCOPY /Y %MSN_TEST_WORKSPACE%\Output\%%i\RH850_X2x*TestLog.log .\
    )
	
    :: Commit Log file
    svn commit -m "Commit via Jenkins - Test Log"
    CD /d %mypath%
)


:: delete dlls at location D:\test\external\X2x\common\generic\generator
DEL /S /Q "U:\external\X2x\common\generic\generator\*.dll"
DEL /S /Q %MSN%_running.txt

GOTO EXIT

:WAIT
sleep 10
goto CHECK

:EXIT
IF NOT "%RESULT%"=="" (
	@ECHO %RESULT%
)
GOTO END

:END
