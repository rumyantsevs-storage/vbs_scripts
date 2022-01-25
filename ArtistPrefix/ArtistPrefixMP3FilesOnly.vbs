Dim objFSO
Dim objFolder
Dim sPrefix
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder(".")
sPrefix = objFolder.Name + " - "
For Each SongFile In objFolder.Files
	If InStr(SongFile.Name, sPrefix) = 0 Then
		If InStr(SongFile.Name, ".vbs") = 0 Then
			If (InStr(SongFile.Name, ".mp3") <> 0 Or InStr(SongFile.Name, ".MP3") <> 0) Then
				objFSO.MoveFile SongFile.Name, sPrefix + SongFile.Name
			End If
		End If
	End If
Next
