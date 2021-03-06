Description
========
The vbs-logging project represents a simplistic logging framework for use in 
Visual Basic scripts. Take the guesswork out of your scripts by logging
pertinent information and errors to file.

The logging framework is implemented in a Logging.vbs script file. We use
a Windows Script File (e.g. LogonScript.wsf) to combine the Logging.vbs file 
with the actual script that does the work (e.g. LogonScript.vbs). We then use 
a batch file to call cscript to execute the LoginScript.wsf. Optionally,
you can skip the batch file and run the script command directly using whatever
means of execution you prefer.

Refer to the "Usage" below for more detailed instructions.

Disclaimer
========

I should mention that I'm not really a VB script guy. I'm a developer that was
tired of people writing poor vb scripts in installation programs as well as the
IT department not having a clue as to why my logon script never worked on 
some machines. Developers and IT professionals must learn that logging is their
friend as it provides an excellent method of performing "in the field" and
"after the fact"  debugging. Stop writing for only the happy case! Stuff is
going to break and you will have to debug it. Be kind to your future self!

Index
========
* examples
	* LogonScript - an example network logon script that makes use of the 
	                logging framework.
		* Logging.vbs - the vb script implementation of the logging framework
		* LogonScript.bat - entry point for executing the logon script
		* LogonScript.wsf - Windows Script File that allows combining the 
		                    Logging.vbs script with the actual LogonScript.vbs
		* LogonScript.vbs - the actual logon script
* src
	* Logging.vbs - the vb script implementation of the logging framework
	* ScriptTemplate.bat - template for creating the entry point for executing
	                       the script
	* ScriptTemplate.vbs - template for the actual script you will create
	* LogonScript.wsf - Windows Script File that allows combining the 
		                Logging.vbs script with the actual LogonScript.vbs

Usage
========
Each file in the source directory contains a decent amount of documentation
on its usage. However, the gist is as follows:

1. Create a copy of the src directory and its files. 
2. Rename the ScriptTemplate.* files to your desired identifier. For example,
   if your are writing a script to clear our temp files you might choose
   ClearTempFiles as your identifier.
3. Within each file, replace all occurences of "ScriptTemplate" with your
   chosen identifier.
4. Edit the <Identifier>.vbs script to perform the actual work you desire. 
   * Typically you will want implement the main functionality in the 
     PerformScriptWork function (altering the argument list as necessary).
   * You will want to adjust command line parameters handling (see 
     "If objArgs.count < 2 Then) and the usage in the DisplayUsage sub

Refer to the example LogonScript for a concrete example.