Dim objFSO
Dim objFolder
Dim sPrefix
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder(".")
sPrefix = objFolder.Name + " - "
For Each AlbumDir In objFolder.SubFolders
	If InStr(AlbumDir.Name, sPrefix) <> 0 Then
		objFSO.MoveFolder AlbumDir.Name, Replace(AlbumDir.Name, sPrefix, "")
	End If
Next
For Each SongFile In objFolder.Files
	If InStr(SongFile.Name, sPrefix) <> 0 Then
		If InStr(SongFile.Name, ".vbs") = 0 Then
			objFSO.MoveFile SongFile.Name, Replace(SongFile.Name, sPrefix, "")
		End If
	End If
Next
