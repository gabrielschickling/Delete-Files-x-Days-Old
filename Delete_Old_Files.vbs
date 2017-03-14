'Written by Gabriel Schickling
'Last Changed : 2/22/2017
'Purpose: To Delete Files that are a certian amount of days old within a specific directory
'Can be ran with Remote Managment or other tools
'Usage: cscript.exe /days: /path: /filetype: (file type is optional)

Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")
'variable to count arguments passed
iNumberOfArguments = WScript.Arguments.Count
'named arguments
Set namedArgs = WScript.Arguments.Named
'Days Argument, would be passed as /days: 36 ( This parameter is required)
strDays = namedArgs.Item("days")
'Path Argument, would be passed as /path: C:\STI
strPath = namedArgs.Item ("path")
'File Type Argument would be passed as /filetype:zip (this parameter is optional)
strfiletype = namedArgs.Item ("filetype")

'checks if days argument is passed, if not quits script
'this is to prevent delete all files in a directory since named argument are optional within vbscript
If Not namedArgs.Exists("days") Then
'quits script if days argument not passed
  Wscript.Quit
'if arguments passed are 2 (I.E) day and path then call search sub process
ElseIf iNumberOfArguments = 2 Then
Call Search ("" & strPath & "")
'if arguments passed are 3 (I.E) day, path and filetype call Searchfiletype sub process
Elseif iNumberOfArguments = 3 Then
Call Searchfiletype ("" & strPath & "")
'if anything else, quit script (hopefully will be able to create a log file to send info back to Remote Managment in the future)
Else
   WScript.Quit
End If

Sub Search(str)
    Dim objFolder, objSubFolder, objFile
    'obj folder is equal to previously passed paramiter
    Set objFolder = objFSO.GetFolder(str)
    For Each objFile In objFolder.Files
        'If file is greater then or equal to amount of days old then delete
        If objFile.DateLastModified <= (Now() - strDays) Then
          'Deletes Selected Files
          objFile.Delete(True)
        End If
    Next
End Sub

Sub Searchfiletype(str)
    Dim objFolder, objSubFolder, objFile
    'obj folder is equal to previously passed paramiter
    Set objFolder = objFSO.GetFolder(str)
    For Each objFile In objFolder.Files
      'if file in the selected path matches paramter passed by filetype arugment then it will go on to next if statement to compare date
      If LCase(Right(Cstr(objFile.Name), 3)) = ("" & strfiletype & "") Then
                     'If file is greater then or equal to amount of days old then delete
                      If objFile.DateLastModified <= (Now() - strDays) Then
                     'Deletes Selected Files
                      objFile.Delete(True)
                      End If
      End If
    Next
End Sub
'Quits Script (Log info here for the future)'
WScript.Quit
