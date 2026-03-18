' VBScript to recursively list all files with relative paths,
' write to mod_files.txt in UTF-8 format, each line prefixed with "add ".
' Excluded files (at root only): mod_files_AddAll.vbs, mod_files.txt, mod_version.ini, changelog.md
' Runs silently, no output.

Option Explicit

Dim fso, rootFolder, excludeList, outputFile, stream

' Create FileSystemObject
Set fso = CreateObject("Scripting.FileSystemObject")

' Get the current directory (script location)
rootFolder = fso.GetAbsolutePathName(".")

' Define the list of files to exclude (exact filenames, case-insensitive comparison)
excludeList = Array("mod_files_AddAll.vbs", "mod_files.txt", "mod_version.ini", "changelog.md")

' Output file path
outputFile = fso.BuildPath(rootFolder, "mod_files.txt")

' Create ADODB.Stream for UTF-8 writing
Set stream = CreateObject("ADODB.Stream")
stream.Type = 2                 ' adTypeText
stream.Charset = "utf-8"        ' UTF-8 without BOM
stream.Open

' Start recursive traversal from root folder with empty relative prefix
TraverseFolder rootFolder, ""

' Save the stream to file (overwrite if exists)
stream.SaveToFile outputFile, 2 ' adSaveCreateOverWrite
stream.Close

Set stream = Nothing
Set fso = Nothing

' Recursive subroutine to traverse folders and write file entries
Sub TraverseFolder(folderPath, relativePrefix)
    Dim folder, file, subFolder, fileName, relPath, excludeName, isExcluded
    
    Set folder = fso.GetFolder(folderPath)
    
    ' Process all files in current folder
    For Each file In folder.Files
        ' Build relative path: if relativePrefix is empty, use just file name; otherwise prefix + "/" + file name
        If relativePrefix = "" Then
            relPath = file.Name
        Else
            relPath = relativePrefix & "/" & file.Name
        End If
        
        ' Check if this file should be excluded (only when at root folder)
        isExcluded = False
        If folderPath = rootFolder Then
            fileName = file.Name
            For Each excludeName In excludeList
                If LCase(fileName) = LCase(excludeName) Then
                    isExcluded = True
                    Exit For
                End If
            Next
        End If
        
        ' Write line if not excluded
        If Not isExcluded Then
            stream.WriteText "add " & relPath & vbCrLf
        End If
    Next
    
    ' Recurse into subfolders
    For Each subFolder In folder.SubFolders
        ' Update relative prefix: if empty, use subfolder name; otherwise prefix + "/" + subfolder name
        If relativePrefix = "" Then
            TraverseFolder subFolder.Path, subFolder.Name
        Else
            TraverseFolder subFolder.Path, relativePrefix & "/" & subFolder.Name
        End If
    Next
End Sub
