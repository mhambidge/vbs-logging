'==================================================================================================
'= Purpose: A general purpose logging library for use within other VB scripts.
'= Usage: 
'=  1. Create a Windows Script File (.wsf) that includes your script and the this logging script.
'=     An example contents of the .wsf file are as follows:
'=       <job id="DetectSQLExp">
'=         <script language="VBScript" src="Logging.vbs"/>
'=         <script language="VBScript" src="DetectSQLExp.vbs"/>
'=       </job>
'=  2. Call OpenLogFile("<file>") near the beginning of the script.
'=  3. Throughout the script log messages of the desired level using the Log<Level>() Subs.
'=  4. Call CloseLogFile() near the end of the script or use the ExitScript method.
'=  5. Run your script using the .wsf file.
'=  
'=  Refer to the TestLogFile() Sub below for example usage.
'=
'==================================================================================================

Option Explicit

Const OPEN_FILE_FOR_READING   = 1
Const OPEN_FILE_FOR_WRITING   = 2
Const OPEN_FILE_FOR_APPENDING = 8

' The log levels. These are hierarchical in nature. If the active log level
' is set to LOG_LEVEL_INFO then messages at LOG_LEVEL_TRACE and LOG_LEVEL_DEBUG
' will be excluded from logging.
Const LOG_LEVEL_TRACE = 0
Const LOG_LEVEL_DEBUG = 1
Const LOG_LEVEL_INFO  = 2
Const LOG_LEVEL_WARN  = 3
Const LOG_LEVEL_ERROR = 4
Const LOG_LEVEL_FATAL = 5

' The delimiter used between fields in the log message
Const DEFAULT_LOG_DELIMITER = " | "

' If enabled log messages are also output to the console
Const DEFAULT_LOG_ECHO_ENABLED = false ' don't enable this when in production

' The default log level.
Const DEFAULT_LOG_LEVEL = 0

Dim oLogFile        ' The log file
Dim oLogFSO         ' The log file file system object
Dim ilogLevel       ' The active log level
Dim sLogID          ' An ID to output in log messages
Dim blogEchoEnabled ' Whether or not to echo log messages to the console
Dim slogDelimiter   ' Delimiter used between sections of the log message

' Perform some initialization of defaults
SetLogLevel DEFAULT_LOG_LEVEL
SetLogEchoEnabled DEFAULT_LOG_ECHO_ENABLED
SetLogDelimiter DEFAULT_LOG_DELIMITER
Set oLogFile = Nothing
Set oLogFSO = Nothing
sLogID = Empty

' Uncomment the following line to run the unit test of the logging functionality
'TestLogFile

'
' Opens the specified log file. Subsequent calls to log messages will log to this file
'
Sub OpenLogFile(pLogFileName)
	Set oLogFSO = CreateObject("Scripting.FileSystemObject")
	Set oLogFile = oLogFSO.OpenTextFile(pLogFileName, OPEN_FILE_FOR_APPENDING, True)
End Sub

'
' Closes the log file, if one is open.
'
Sub CloseLogFile()
	If (oLogFile Is Nothing) Then
		Exit Sub
	End If
	
	oLogFile.Close
	Set oLogFile = Nothing
	Set oLogFSO = Nothing
End Sub

'
' Logs the success of failure of an operation. Err.Number is used to test success
' or failure. 
'
' Arguments:
' - sOperation - a textual description of the operation.
'
Sub LogOperation(sOperation) 
	LogOperation_sub sOperation, false, 0
End Sub

'
' Logs the success of failure of an operation. Err.Number is used to test success
' or failure. 
'
' Arguments:
' - sOperation - a textual description of the operation.
'
Sub LogOperationAndClear(sOperation) 
	LogOperation_sub sOperation, false, 0
	Err.Clear
End Sub

'
' Logs the success of failure of an operation. Err.Number is used to test success
' or failure. On failure the script will exit with the provided error code.
'
' Arguments:
' - sOperation - a textual description of the operation.
' - iErrorCode - a custom error code to include in the message if a failure 
'     occurs.
'
Sub LogRequiredOperation(sOperation, iErrorCode)
	LogOperation_sub sOperation, true, iErrorCode
End Sub

'
' Logs the success of failure of an operation. Err.Number is used to test success
' or failure. 
'
' Arguments:
'   - sOperation - a textual description of the operation.
'   - bQuitOnError - if true and the operation failed, quit the script.
'   - iErrorCode - a custom error code to include in the message if a failure 
'       occurs.
'
Sub LogOperation_sub(sOperation, bQuitOnError, iErrorCode)

	If (Err.Number <> 0) Then
		LogError "Operation """ & sOperation & """ failed. " & Err.Number & " - " & Err.Description
		If (bQuitOnError = true) Then
			WScript.Quit(iErrorCode)
		End If
	Else
		LogInfo "Operation """ & sOperation & """ succeeded."
	End If 
	
End Sub

'
' Sets the active log level. Messages with a level lower than this level will
' not be logged.
'
' Arguments:
'   - sLevel - the desired log level, such as LOG_LEVEL_INFO
'
Sub SetLogLevel(sLevel)
	iLogLevel = sLevel
End Sub

' Enabled echoing of log messages to the console.
'
' Arguments:
'   - pLogEchoEnabled - true if log messages should be echoed to the console
'
Sub SetLogEchoEnabled(bNewLogEchoEnabled)
	blogEchoEnabled = bNewLogEchoEnabled
End Sub

' Sets the delimiter to use in log messages
'
' Arguments:
'   - pLogDelimiter - the delimiter to use in log messages
'
Sub SetLogDelimiter(sNewLogDelimiter)
	slogDelimiter = sNewLogDelimiter
End Sub

'
' If set, this idenfying string will be included in each log message. Can be
' useful to identify scripts of functions when logging to the same file.
'
' Arguments:
'   - sNewLogID - the log id to include.
'
Sub SetLogID(sNewLogID)
	sLogID = sNewLogID
End Sub

'
' Logs the specified message at trace level (LOG_LEVEL_TRACE)
'
' Arguments:
'   - sMessage - the message to log.
'
Sub LogTrace(sMessage)
	LogMessage LOG_LEVEL_TRACE, sMessage
End Sub

'
' Logs the specified message at debug level (LOG_LEVEL_DEBUG)
'
' Arguments:
'   - sMessage - the message to log.
'
Sub LogDebug(sMessage)
	LogMessage LOG_LEVEL_DEBUG, sMessage
End Sub

'
' Logs the specified message at info level (LOG_LEVEL_INFO)
'
' Arguments:
'   - sMessage - the message to log.
'
Sub LogInfo(sMessage)
	LogMessage LOG_LEVEL_INFO, sMessage
End Sub

'
' Logs the specified message at warn level (LOG_LEVEL_WARN)
'
' Arguments:
'   - sMessage - the message to log.
'
Sub LogWarn(sMessage)
	LogMessage LOG_LEVEL_WARN, sMessage
End Sub

'
' Logs the specified message at error level (LOG_LEVEL_ERROR)
'
' Arguments:
'   - sMessage - the message to log.
'
Sub LogError(sMessage)
	LogMessage LOG_LEVEL_ERROR, sMessage
End Sub

'
' Logs the specified message at fatal level (LOG_LEVEL_FATAL)
'
' Arguments:
'   - sMessage - the message to log.
'
Sub LogFatal(sMessage)
	LogMessage LOG_LEVEL_FATAL, sMessage
End Sub

'
' Logs the specified message at the specified level. If the active log level is
' higher than the specified level then the message will not be logged.
'
' Arguments:
'   - sLevel - the level at which to log the message.
'   - sMessage - the message to log
'
Sub LogMessage(sLevel, sMessage)
	If sLevel < iLogLevel Then
		Exit Sub
	End If

	Dim logStr
	Dim logLevelString
	logLevelString = GetLogLevelString(sLevel)
	
	logStr = Now
	If (Not IsEmpty(sLogID)) Then
		logStr = logStr & slogDelimiter & sLogID
  End If
	logStr = logStr & slogDelimiter & logLevelString & slogDelimiter & sMessage
	
	If (Not oLogFile Is Nothing) Then
		oLogFile.WriteLine(logStr)
	End If

	If blogEchoEnabled = true Then
		WScript.Echo logStr
	End If
End Sub

'
' Returns the textual representation of the specified log level.
'
' Arguments: 
'  - sLevel - the log level
'
Function GetLogLevelString(sLevel)
	Select Case sLevel
		Case LOG_LEVEL_TRACE
			GetLogLevelString = "TRACE"
		Case LOG_LEVEL_DEBUG
			GetLogLevelString = "DEBUG"
		Case LOG_LEVEL_INFO
			GetLogLevelString = "INFO"
		Case LOG_LEVEL_WARN
			GetLogLevelString = "WARN"
		Case LOG_LEVEL_ERROR
			GetLogLevelString = "ERROR"
		Case LOG_LEVEL_FATAL
			GetLogLevelString = "FATAL"
		Case Else
			GetLogLevelString = "<Unknown>"
	End Select
End Function

'
' Utility function that will log the exit code, close the log file, and exit
' the active script.
'
' Arguments:
'   - iExitCode - the script exit code.
'
Private Sub ExitScript(iExitCode)
	LogDebug(WScript.ScriptName & " exiting with exit code: " & iExitCode)
	CloseLogFile()
	WScript.Quit(iExitCode)
End Sub

'
' Test function
'
Sub TestLogFile
	OpenLogFile("C:\\test.log")
	LogTrace "This is a TRACE level message"
	LogDebug "This is a DEBUG level message"
	LogInfo "This is a INFO level message"
	LogWarn "This is a WARN level message"
	LogError "This is a ERROR level message"
	LogFatal "This is a FATAL level message"
	LogMessage LOG_LEVEL_INFO, "This is a test"
	CloseLogFile()
End Sub
