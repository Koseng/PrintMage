Dim shell,command, arg1, arg2, arg3, file, scriptDir
scriptdir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
file = " -File """ & scriptdir & "\printMage.ps1"""
arg1 = " -pathWorkDir """ & scriptdir & """"
arg2 = " -pdfFullPath """ & WScript.Arguments(0) & """"
arg3 = " -pdfFileBaseName """ & WScript.Arguments(1) & """"
command = "powershell.exe -NoLogo -ExecutionPolicy bypass -NonInteractive" & file & arg1 & arg2 & arg3
'Wscript.Echo command
Set shell = CreateObject("WScript.Shell")
shell.Run command,0
