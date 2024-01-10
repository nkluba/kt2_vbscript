Option Explicit

Dim fso, startFolder, targetFolder, folderCount
Set fso = CreateObject("Scripting.FileSystemObject")

startFolder = "C:\Users\luban\Pictures\Screenshots"
targetFolder = "C:\Users\luban\Pictures\Sorted"

SortFiles startFolder, targetFolder

WScript.Echo folderCount & " folders were created."

Sub SortFiles(folderPath, targetPath)
    Dim folder, file, fileDate, targetDateFolder, subFolder
    Set folder = fso.GetFolder(folderPath)

    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file)) = "jpg" Or LCase(fso.GetExtensionName(file)) = "jpeg" Or LCase(fso.GetExtensionName(file)) = "png" Then
            fileDate = Year(file.DateLastModified) & "\" & Year(file.DateLastModified) & "-" & Right("0" & Month(file.DateLastModified), 2) & "-" & Right("0" & Day(file.DateLastModified), 2)
            targetDateFolder = targetPath & "\" & fileDate
            If Not fso.FolderExists(targetPath & "\" & Year(file.DateLastModified)) Then 
			    fso.CreateFolder(targetPath & "\" & Year(file.DateLastModified))
				folderCount = folderCount + 1
			End If
			
			If Not fso.FolderExists(targetDateFolder) Then 
                fso.CreateFolder(targetDateFolder)
                folderCount = folderCount + 1
            End If
                
            fso.MoveFile file.Path, targetDateFolder & "\"
        End If
    Next

    If folder.SubFolders.Count > 0 Then
        For Each subFolder In folder.SubFolders
            SortFiles subFolder.Path, targetPath
        Next
    End If
End Sub
