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

' Obtain the command line arguments
Set objArgs = WScript.Arguments

' If we didn't get the expected number of command line arguments, display the usage
' and exit
If objArgs.count < 1 Then
	DisplayUsage()
	WScript.Quit(EXIT_INVALID_COMMAND_LINE_ARGUMENTS)
End If

' Otherwise, load the command line arguments into appropriately named variables
sLogFilename = objArgs.item(0)

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
PerformScriptWork()

' If errors happen in the PerformScriptWork function then you can call ExitScript
' from PerformScriptWork to return an error code to the caller. Otherwise we
' fall out with an EXIT_OK result.
ExitScript(EXIT_OK)


' The following are the various functions used in the main body of the script above.

' Perform the actual main work for the script.
Function PerformScriptWork()
	On Error Resume Next

	Dim objAD, objUser, objNetwork, CurrentUserName, strremotepath, WshShell, strUserDN

	'Path to general DFS
	strremotepath = "\\forest.intranet.hambo\UserFiles"
	LogInfo("Using top level DFS path: """ & strremotepath & """")
	
	Set objNetwork = CreateObject("Wscript.Network")
	CurrentUserName = objNetwork.UserName
	LogInfo("Username: " & CurrentUserName)
		
	'Remove current network mappings
	objNetwork.RemoveNetworkDrive "X:", True, True
	LogOperationAndClear("Removed network drive mapping for X:")
	objNetwork.RemoveNetworkDrive "Y:", True, True
	LogOperationAndClear("Removed network drive mapping for Y:")
	objNetwork.RemoveNetworkDrive "Z:", True, True
	LogOperationAndClear("Removed network drive mapping for Z:")

	'Map all access network drives
	objNetwork.MapNetworkDrive "X:", "\\forest.intranet.hambo\Private\users$\" & CurrentUserName
	LogOperationAndClear("Added network drive mapping for X:")
	objNetwork.MapNetworkDrive "Y:", strremotepath & "\BackedUp"
	LogOperationAndClear("Added network drive mapping for Y:")
	objNetwork.MapNetworkDrive "Z:", strremotepath & "\NotBackedUp"
	LogOperationAndClear("Added network drive mapping for Z:")

	'Retrieve User Object from Directory
	Dim Winntuser
	Set Winntuser = GetObject("WinNT://" & objNetwork.UserDomain & "/" & objNetwork.UserName & ",user")
	LogOperationAndClear("Retrieved WinNT User for """ & objNetwork.UserDomain & "/" & objNetwork.UserName & """")

	'Map Engineering drives only if user is a member of those groups
	LogInfo("Handling mapped drives for ""Engineering"" group")
	If IsMemberOfGroup(objNetwork.UserDomain, Winntuser, "Engineering") = True Then
		
		objNetwork.RemoveNetworkDrive "V:", True, True
		LogOperationAndClear("Removed network drive mapping for V:")
		objNetwork.MapNetworkDrive "V:", "\\forest.intranet.hambo\Private\Engineering$"
		LogOperationAndClear("Added network drive mapping for V:")
		
		objNetwork.RemoveNetworkDrive "W:", True, True
		LogOperationAndClear("Removed network drive mapping for W:")
		objNetwork.MapNetworkDrive "W:", "\\forest.intranet.hambo\Private\Builds$"
		LogOperationAndClear("Added network drive mapping for W:")
		
	End If
		
	'Update Group policy on system
	Set WshShell = WScript.CreateObject("WScript.Shell")
	wshshell.run("gpupdate /force"), 0
	LogOperationAndClear("Ran gpupdate")

	' Always reset the error handling back to the default when you are done
	On Error Goto 0

End Function

Function IsMemberOfGroup(Domain, User, Group)
	'Initially, user is not a member, unless proven otherwise
	IsMemberOfGroup = False

	'Continue processing even if errors occur
	On Error Resume Next
	'Retrieve group from Directory
	Dim GroupPath
	Set GroupPath = GetObject("WinNT://" & Domain & "/" & Group & ",group")

	'If group could not be found, display a message
	'otherwise check directory to see if user is a member of the group specified
	If Err.Number Then
		IsMemberOfGroup = "Error"
		'MsgBox "There was no group found called " & Group		'For debugging only
	Else
		IsMemberOfGroup = GroupPath.IsMember(User.ADsPath)
	End If
	
End Function

' Display the script usage
Private Sub DisplayUsage()

	' Turn on echoing to the console since we may not have a log file opened at this point
	SetLogEchoEnabled(true)

	' Display the actual script usage
	LogInfo("Usage: ")
	LogInfo("    " & WScript.ScriptName & " ""<Log Filename>""")

End Sub