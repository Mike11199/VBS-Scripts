'This is a generic macro/VBS script for copying a file from a network drive into the same folder as the VBS script

'This does NOT contain any confidential company information

'Replace the folder name, and string, extension if using this script, and run from a .vbs file from the folder you want to copy a file into

'This is a convenient for a PDF file that is generated, where one wants to copy the newest file into the folder.  It will also copy files with the same name matching
'that string created within the past few days, to ensure no files are missed.

dim oFSO
dim oShell
dim oFolder
dim oFile 

dim sNewest
dim sFile
dim sFolder

Set oFSO = CreateObject("Scripting.FileSystemObject")
Set FolderRef = oFSO.GetFolder("c:\")

'============CHANGE THIS TO CHANGE SOURCE==========================================

Set FolderRef = oFSO.GetFolder("REPLACE WITH FOLDER")  'put the folder name you want to start in here

dim folderString
Dim fso
Dim curDir
dim lastModified
Dim WinScriptHost
Set fso = CreateObject("Scripting.FileSystemObject")
curDir = fso.GetAbsolutePathName(".")
curDir = curDir & "\"
Set fso = Nothing
dim searchFileName
searchFileName = "REPLACE WITH FILE STRING"  'search file name string to search for here (can be part of file name, not entire match)


Set oShell = CreateObject("WScript.Shell")


'modify this to change date
threeDaysBefore = FormatDateTime(Now-7, 2)


'========THIS LOOPS THROUGH EACH FILE AND SELECTS THE NEWEST ONE===================

For Each sFile In FolderRef.Files
lastModified = sFile.DateLastModified
lastModified = FormatDateTime(lastModified , 2)
If  DateDiff("d",lastModified,threeDaysBefore ) < 7 Then	
               If instr(sFile.name, searchFileName) <> 0 Then
              	 Call oFSO.CopyFile(sFile, curDir, True)
               End If
End If
Next




'===========THIS SECTIONS FIND THE FILE COPIED IN THE CURRENT FOLDER AND OPENS IT=====
dim folder
dim file
dim copiedfile

dim folderName
dim extension

extension = "PDF"   'change if file not PDF


curDir = curDir & "\"

Set FolderRef = oFSO.GetFolder(curDir )

' Loop over all files in the folder until the searchFileName is found
For each file In FolderRef.Files    

    If instr(file.name, searchFileName) <> 0 Then
   	 If instr(file.name, extension) <> 0 Then
        
          copiedfile = file
        ' Exit the loop, we only want to rename one file
        Exit For
    End If
End If
Next

'msgbox copiedfile
Set WshShell = CreateObject("WScript.Shell")
WshShell.Run """" & copiedfile & """"

WshShell.SendKeys "^+{R}"
