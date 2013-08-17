@ECHO OFF

REM This batch file will call the the Windows Script Host to run the 
REM ScriptTemplate Windows Script File (.wsf).
REM
REM The "%*" means pass all the command line arguments provided to this
REM batch file on through to the script.
cscript LogonScript.wsf %*