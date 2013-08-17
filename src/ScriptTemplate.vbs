'==================================================================================================
'= Purpose: Sample script to demonstrate the use of the Logging.vbs logging framework. This script 
'= should be called using the ScriptTemplate.wsf file rather than using this vbs file.
'=
'= Usage: See DisplayUsage() Sub below
'= Exit Codes: See EXIT_* constants defined below.
'=
'==================================================================================================

' Option Explicit requires that variabled be properly defined. We insist on using this to
' prevent lazy coders.
Option Explicit

' By default we want the script to die when it encounters an error. Otherwise you end up having
' hard to find bugs.
'On Error Resume Next
On Error Goto 0

' Define the various exit codes for the script then use the ExitScript
' function to return these to the caller, e.g. ExitScript(EXIT_GENERAL_FAILURE)
Const EXIT_OK                             = 0
Const EXIT_INVALID_COMMAND_LINE_ARGUMENTS = -1
Const EXIT_GENERAL_FAILURE                = -2

' Define the variables we will be using to process the command line arguments
Dim objArgs
Dim sLogFilename
Dim sUserMessage

' Obtain the command line arguments
Set objArgs = WScript.Arguments

' If we didn't get the expected number of command line arguments, display the usage
' and exit
If objArgs.count < 2 Then
	DisplayUsage()
	WScript.Quit(EXIT_INVALID_COMMAND_LINE_ARGUMENTS)
End If

' Otherwise, load the command line arguments into appropriately named variables
sLogFilename = objArgs.item(0)
sUserMessage = objArgs.item(1)

' Specify the file to log to using the OpenLogFile method.
'
' The logging functionality will append to an existing file. If you don't want 
' this then delete the file if it already exists before opening the log file.
' 
' Note that you don't have to specify a log file and can instead only log to 
' the console using SetLogEchoEnabled(true), discussed below
OpenLogFile(sLogFilename)

' Use the SetLogID to provide a string that will be included in log message to 
' identify log messages from this script. Useful if multiple scripts will be logging
' to the same file. (Note that this doesn't mean the logging supports multiple scripts
' writing to it at the same time. Rather, this is if you run scripts serially and want
' them to all log to the same file.
SetLogID(WSCript.ScriptName)

' By default the log level is set to TRACE, so you will see all log messages.
' You can use the SetLogLevel method to change the level of messages you want
' output
SetLogLevel(LOG_LEVEL_TRACE)

' By default the logging functionality only writes to the log file. To have
' it also write to the console use SetLogEchoEnabled(true)
SetLogEchoEnabled(true)

' If you want to use a delimiter other than the pipe symbol, you can do so 
' using the SetLogDelimiter method
'SetLogDelimiter(" + ")

' Perform the actual work for the script in the PerformScriptWork method (you can
' name this whatever you want). Pass in the required command line data.
PerformScriptWork(sUserMessage)

' If errors happen in the PerformScriptWork function then you can call ExitScript
' from PerformScriptWork to return an error code to the caller. Otherwise we
' fall out with an EXIT_OK result.
ExitScript(EXIT_OK)


' The following are the various functions used in the main body of the script above.

' Perform the actual main work for the script.
Function PerformScriptWork(sUserMessage)
	On Error Resume Next

	' The following is an example of using the logging functionality. 

	' You can use the various utility methods to log at different log levels.
	LogTrace "This is a TRACE level message"
	LogDebug "This is a DEBUG level message"
	LogInfo "This is a INFO level message"
	LogWarn "This is a WARN level message"
	LogError "This is a ERROR level message"
	LogFatal "This is a FATAL level message"

	' Alternatively you can call the main logging method directly and pass the desired
	' log level
	LogMessage LOG_LEVEL_INFO, "This is a test"

	' For grins we log the message provided by the user on the command line
	LogInfo("Here is the user message provided on the command line: " & sUserMessage)
	
	
	' The logging framework also provides methods to check the Err.Number to determine
	' if an operation failed. You need to use On Error Resume Next to allow checking
	' of the error code, for example
	On Error Resume Next
	' Do something here that can generate an error. We'll try to create a nonsense object
	Set oLogFSO = CreateObject("ThisDoesNotExist")
	' Now we call the LogOperation method which will check the error code for us and log
	' either an INFO message on success or an ERROR message on failure.
	LogOperation("Creating a fake object")
	' Note that we could also call LogOperationRequired("Creating a fake object", EXIT_GENERAL_ERROR)
	' which has the additional effect that on error it will exit the script with the provided error code.

	' Always reset the error handling back to the default when you are done
	On Error Goto 0

End Function

' Display the script usage
Private Sub DisplayUsage()

	' Turn on echoing to the console since we may not have a log file opened at this point
	SetLogEchoEnabled(true)

	' Display the actual script usage
	LogInfo("Usage: ")
	LogInfo("    " & WScript.ScriptName & " ""<Log Filename>"" ""<Some Random Text>""")

End Sub