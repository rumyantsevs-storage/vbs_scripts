' ****************************************************************************
' *                          Custom User Variables                           *
' ****************************************************************************
Dim sTargetPath
Dim sArguments
Dim sWorkingDirectory
sTargetPath = "C:\Program Files\Opera\66.0.3515.103\opera.exe"
sWorkingDirectory = "C:\Program Files\Opera\66.0.3515.103"
sArguments = Array( _
				   "do", "www.discogs.com", _
				   "gc", "www.google.com", _
				   "gi", "www.google.com/imghp", _
				   "gt", "translate.google.com/#en/ru", _
				   "ov", "www.oldversion.com", _
				   "rt", "www.rutracker.org", _
				   "sf", "www.savefrom.net", _
				   "we", "en.wikipedia.org", _
				   "wr", "ru.wikipedia.org", _
				   "ya", "www.ya.ru", _
				   "yi", "www.yandex.ru/images", _
				   "ym", "mail.yandex.ru", _
				   "yt", "www.youtube.com" _
				  )
' ****************************************************************************

Set WshShell = WScript.CreateObject("WScript.Shell")

For i = 0 To UBound(sArguments) Step 2
	Set Lnk = WshShell.CreateShortcut("C:\Windows\system\" + sArguments(i) + _
									  ".lnk")
	Lnk.TargetPath = sTargetPath
	Lnk.Arguments = sArguments(i + 1)
	Lnk.WorkingDirectory = sWorkingDirectory
	Lnk.Save
Next
