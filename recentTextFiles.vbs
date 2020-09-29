' Recent Text Files List
'  by Bill Lugo
'
' Retrieves list of recently-opened text files.
' Environment: Windows 10

' initialization
outputFile = "tempRecentTextFile.txt"

Set objShell = CreateObject("Wscript.Shell")
objShell.CurrentDirectory = objShell.ExpandEnvironmentStrings("%AppData%") + "\Microsoft\Windows\Recent"

tempFile = objShell.ExpandEnvironmentStrings("%AppData%") + "\Microsoft\Windows\Recent\" + outputFile
' %temp% would be good, but there's a permissions issue involved


' Get list of text files
objShell.Run("cmd /C dir /b /o-d *.txt.lnk>" + tempFile)


' Open file list
set filesys = CreateObject("Scripting.filesystemObject")
fileList = CreateObject("Scripting.FileSystemObject").openTextFile(tempFile).readAll()


' Clean up - delete temp file
If filesys.FileExists(tempFile) Then
 filesys.deleteFile(tempFile)
End If


' Output
output = "Recent files: " + chr(13) + chr(13) + fileList
MsgBox output,,"Recent text files"

'TODO: create a form with clickable links to open each file