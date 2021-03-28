@ECHO OFF

@REM OBS.
@REM For at funktionen skal virke ma man ikke gore noget imens scriptet korer
@REM - skriver DONE nar det er ferdigt

@REM Specificer hvilken mappe mappen/objektet ligger i. Dernest mappen/objektet som du have effectiveaccess pa,
@REM Også navnet på brugeren/objektet (PS. er kun testet til at virke med brugere)
@REM -
@REM effectiveaccess.cmd "<stien til mappen>" "<Navnet pa mappen>" "<navnet pa brugeren>"
@REM f.eks. - bemærk at parameterne skal pakkes ind i gase ojne.
@REM effectiveaccess.cmd "D:\Firm" "Marketing" "%USERNAME%" "Marketing_%USERNAME%"
@REM effectiveaccess.cmd "C:\Users" "administrator" "%USERNAME%" "administrator_%USERNAME%"

SET _PATHH=%1
SET _FOLDERR=%2
SET _USERNAMEE=%3
SET _PICTUREE=%3

@REM fjerner gase ojne 1/2
@REM Har fundet det her stykke kode som fjerner gaseojne.
@REM Det her er en darlig implementering (ulaeselig) - burde vaere i ejen fil for sig selv som i vejledningen
@REM kilde https://ss64.com/nt/syntax-dequote.html
@REM Fjerner gaseojne ("") fra parameterne
@REM ECHO %_PATHH% %_FOLDERR% %_USERNAMEE%
CALL :dequote _PATHH
CALL :dequote _FOLDERR
CALL :dequote _USERNAMEE
CALL :dequote _PICTUREE

@REM 'Abner explorer
explorer
ping -n 2 127.0.0.1 > nul
:VBSDynamicBuild
SET TempVBSFile=%tmp%\~tmpSendKeysTemp.vbs
IF EXIST "%TempVBSFile%" DEL /F /Q "%TempVBSFile%"
@REM  ctrl + l
ECHO Set WshShell = WScript.CreateObject("WScript.Shell")   >>"%TempVBSFile%"
ECHO Wscript.Sleep 200                                      >>"%TempVBSFile%"
ECHO WshShell.SendKeys "^l"                                 >>"%TempVBSFile%"
ECHO Wscript.Sleep 200                                      >>"%TempVBSFile%"
@REM Paster stien til mappen ind i file explore
ECHO WshShell.SendKeys "%_PATHH%"                                 >>"%TempVBSFile%"
ECHO Wscript.Sleep 200                                      >>"%TempVBSFile%"
ECHO WshShell.SendKeys "{ENTER}"                                 >>"%TempVBSFile%"
ECHO Wscript.Sleep 200                                      >>"%TempVBSFile%"
@REM Skriver navnet pa mappen for at target pa den
ECHO WshShell.SendKeys "%_FOLDERR%"                                 >>"%TempVBSFile%"
ECHO Wscript.Sleep 200                                      >>"%TempVBSFile%"
@REM Abner Propeties
ECHO WshShell.SendKeys "%%{ENTER}"                                 >>"%TempVBSFile%"
ECHO Wscript.Sleep 200                                      >>"%TempVBSFile%"
ECHO WshShell.SendKeys "{TAB 5}"                                 >>"%TempVBSFile%"
ECHO Wscript.Sleep 200                                      >>"%TempVBSFile%"
@REM Navigere til security tappen
ECHO WshShell.SendKeys "{RIGHT 2}"                                 >>"%TempVBSFile%"
ECHO Wscript.Sleep 200                                          >>"%TempVBSFile%"
ECHO WshShell.SendKeys "{TAB 10}"                                 >>"%TempVBSFile%"
ECHO Wscript.Sleep 200                                          >>"%TempVBSFile%"
@REM Abner Advanced security
ECHO WshShell.SendKeys "{ENTER}"                                 >>"%TempVBSFile%" 
ECHO Wscript.Sleep 200                                          >>"%TempVBSFile%"
@REM @REM Maksimere vindue 
ECHO WshShell.SendKeys "%% "                                    >>"%TempVBSFile%"
ECHO Wscript.Sleep 200                                          >>"%TempVBSFile%"
ECHO WshShell.SendKeys "{DOWN 4}"                                   >>"%TempVBSFile%"
ECHO Wscript.Sleep 200                                          >>"%TempVBSFile%"
ECHO WshShell.SendKeys "{ENTER}"                                 >>"%TempVBSFile%"
@REM Navigere til effectiveaccess tappen 
ECHO WshShell.SendKeys "{TAB 9}"                                 >>"%TempVBSFile%"
ECHO Wscript.Sleep 200                                      >>"%TempVBSFile%"
ECHO WshShell.SendKeys "{RIGHT}"                                 >>"%TempVBSFile%"
ECHO Wscript.Sleep 200                                          >>"%TempVBSFile%"
ECHO WshShell.SendKeys "{TAB 9}"                                 >>"%TempVBSFile%"
ECHO Wscript.Sleep 200                                      >>"%TempVBSFile%"
ECHO WshShell.SendKeys "{RIGHT}"                                 >>"%TempVBSFile%"
ECHO Wscript.Sleep 200                                          >>"%TempVBSFile%"
ECHO WshShell.SendKeys "{TAB}"                                 >>"%TempVBSFile%"
ECHO Wscript.Sleep 200                                          >>"%TempVBSFile%"
@REM Vaelger en bruger
ECHO WshShell.SendKeys " "                                 >>"%TempVBSFile%"
ECHO Wscript.Sleep 200                                          >>"%TempVBSFile%"
ECHO WshShell.SendKeys "{TAB}"                                 >>"%TempVBSFile%"
ECHO Wscript.Sleep 200                                          >>"%TempVBSFile%"
ECHO WshShell.SendKeys "{ENTER}"                                 >>"%TempVBSFile%"
ECHO Wscript.Sleep 200                                          >>"%TempVBSFile%"
ECHO WshShell.SendKeys "{TAB 5}"                                 >>"%TempVBSFile%"
ECHO Wscript.Sleep 200                                          >>"%TempVBSFile%"
ECHO WshShell.SendKeys "%_USERNAMEE%"                                 >>"%TempVBSFile%"
ECHO Wscript.Sleep 200                                          >>"%TempVBSFile%"
ECHO WshShell.SendKeys "{TAB 2}"                                 >>"%TempVBSFile%"
ECHO Wscript.Sleep 200                                          >>"%TempVBSFile%"
ECHO WshShell.SendKeys "{ENTER}"                                 >>"%TempVBSFile%"
ECHO Wscript.Sleep 200                                          >>"%TempVBSFile%"
ECHO WshShell.SendKeys "{TAB}"                                 >>"%TempVBSFile%"
ECHO Wscript.Sleep 200                                          >>"%TempVBSFile%"
ECHO WshShell.SendKeys "{ENTER}"                                 >>"%TempVBSFile%"
ECHO Wscript.Sleep 200                                          >>"%TempVBSFile%"
ECHO WshShell.SendKeys "{TAB 3}"                                 >>"%TempVBSFile%"
ECHO Wscript.Sleep 200                                          >>"%TempVBSFile%"
ECHO WshShell.SendKeys "{ENTER}"                                 >>"%TempVBSFile%"
ECHO Wscript.Sleep 200                                          >>"%TempVBSFile%"
ECHO WshShell.SendKeys "{TAB 6}"                                 >>"%TempVBSFile%"
ECHO Wscript.Sleep 200                                          >>"%TempVBSFile%"
ECHO WshShell.SendKeys "{ENTER}"                                 >>"%TempVBSFile%"
ECHO Wscript.Sleep 200                                          >>"%TempVBSFile%"
@REM scroller ned i bunder sa man kan se alle rettigheder
ECHO WshShell.SendKeys "{DOWN 16}"                                   >>"%TempVBSFile%"
@REM Tager et screenshot
ECHO WshShell.run ".\nircmd.exe cmdwait 200 savescreenshot %USERPROFILE%\Desktop\%_PICTUREE%.png" >>"%TempVBSFile%"
ECHO Wscript.Sleep 400                                          >>"%TempVBSFile%"
ECHO WshShell.SendKeys "%%{F4}"                                 >>"%TempVBSFile%"
ECHO Wscript.Sleep 200                                          >>"%TempVBSFile%"
ECHO WshShell.SendKeys "%%{F4}"                                 >>"%TempVBSFile%"
ECHO Wscript.Sleep 200                                          >>"%TempVBSFile%"
ECHO WshShell.SendKeys "%%{F4}"                                 >>"%TempVBSFile%"
CSCRIPT //nologo "%TempVBSFile%"

@REM fjerner gase ojne 2/2
@REM Har fundet det her stykke kode som fjerner gaseojne.
@REM Det er en darlig implementering - burde vaere i ejen fil for sig selv som i vejledningen
@REM kilde https://ss64.com/nt/syntax-dequote.html
Goto :eof
:DeQuote
for /f "delims=" %%A in ('echo %%%1%%') do set %1=%%~A
Goto :eof