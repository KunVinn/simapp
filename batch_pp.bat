@ECHO OFF
setlocal EnableDelayedExpansion
:: runElmFor_Noddis Folder_with_elmfor.bin Folder_with_xlsx -50 49 [RUN]
set FLine=2F
:: for %i in (B L)do runElmFor_Noddis0.bat Wisting%i_RifV_ALS\condSet%FLine% Wisting%i_RifV_ALS\condSet%FLine% -50 49 [RUN]
set "pattern=\condSet%FLine%"
set nLC=1
REM for /f "usebackq delims=" %%A in ("lcid!FLine!.txt") do (
for /f "usebackq delims=" %%A in ("lcid%FLine%.txt") do (
    set "LC(!nLC!)=%%A"
    set /a nLC+=1
)

set resN=2
set /a iCount = 0
IF {"%~3"}=={""} ( set /a "nS = -28"
) else set /a "nS=%~3"
REM -28

IF {"%~4"}=={""} ( set /a "nL=-(nS+1)"
) else set /a "nL=%~4"
REM 27
echo.%nS% and %nL%
pause
IF {"%~5"}=={"RUN"} goto :RUN
for /f %%i in ('dir/b/ad/s %~1'
) do (
 set lc=%%i
 if exist !lc!\key_sima_elmfor.txt echo.^^#!iCount! folder: !lc! Exists OK to proceed^^!
 if !iCount! == 0 (robocopy %%i %%i /s /move)
 if exist !lc!\key_sima_elmfor.txt (copy/y key_sima_elmfor.xlsx !lc!\key_sima_elmfor.txt.xlsx&SimaPp /e!lc!\key_sima_elmfor.txt /lelm.txt /bI /d0)
 if exist !lc!\key_sima_noddis.txt (copy/y key_sima_noddis.xlsx !lc!\key_sima_noddis.txt.xlsx&SimaPp /n!lc!\key_sima_noddis.txt /lnod.txt /bI /d0)
 if exist !lc!\results.txt (copy/y results.xlsx !lc!\results.txt.xlsx&SimaPp /r!lc!\results.txt /lres!resN!.txt /bI /d0)
 for /l %%j in (1 1 15) do if exist !lc!_%%j\key_sima_elmfor.txt (copy/y key_sima_elmfor.xlsx !lc!_%%j\key_sima_elmfor.txt.xlsx&SimaPp /e!lc!_%%j\key_sima_elmfor.txt /lelm.txt /bI /d0)
 for /l %%j in (1 1 15) do if exist !lc!_%%j\key_sima_noddis.txt (copy/y key_sima_noddis.xlsx !lc!_%%j\key_sima_noddis.txt.xlsx&SimaPp /n!lc!_%%j\key_sima_noddis.txt /lnod.txt /bI /d0)
 for /l %%j in (1 1 15) do if exist !lc!_%%j\results.txt (copy/y results.xlsx !lc!_%%j\results.txt.xlsx&SimaPp /r!lc!_%%j\results.txt /lres!resN!.txt /bI /d0)
 if exist !lc!\key_sima_elmfor.txt set /a iCount+=1
)

:RUN

rem -- Random number between 1 and 50
rem set /a k=%random% %% 50 + 1
IF NOT {%~2}=={} (
	for /f %%i in ('dir/b/s %~2\*elmfor.txt.xlsx')do (
    set "p=%%~di%%~pi"
    REM set "l=!p:\condSet!FLine!=!"
    set "l=!p:%pattern%=!"
    REM set "n=!p:*\condSet!FLine!_=!"
    set "n=!p:*%pattern%_=!"
    set "k=!n:\=!"
    rem -- Use helper variable to retrieve the value
    set "index=!k!"
    for %%M in (!index!) do set "value=!LC(%%M)!"
    echo LC!k!=!value!
    ModelTest2Xls /p!p! "/l!l:~%nS%,%nL%!_!value!" /xResSum.xlsx "/tSIMA SIMAPlt" /sX /MBL=19748 /c1000.
    ::ModelTest2Xls /p!p! "/l!l:~%nS%,%nL%!_!value!" /xResSum.xlsx "/tSIMA SIMAPlt" /sX /MObs /MBL=19748 /c1000.
    ::ModelTest2Xls /p!p! "/l!l:~%nS%,%nL%!_!value!" /xResSum.xlsx "/tSIMA SIMAPlt" /sX /MWFT /MBL=19748 /c1000.
	)
)
goto :EOF


