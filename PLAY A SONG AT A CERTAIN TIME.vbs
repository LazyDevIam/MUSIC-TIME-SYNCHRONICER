Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

strMP3Path = InputBox("Enter the path of the MP3 file to play:")
Do While Not objFSO.FileExists(strMP3Path)
    strMP3Path = InputBox("The file does not exist. Please enter a valid path to the MP3 file to play:")
    If strMP3Path = "" Then
        WScript.Quit
    End If
Loop

strTime = InputBox("Enter the time to play the MP3 file (format: HH:MM:SS AM/PM):")
dtDateTime = CDate(Date() & " " & strTime)

strConfirm = MsgBox("MP3 file will be played at " & FormatDateTime(dtDateTime, vbLongTime) & ". Do you want to proceed?", vbYesNo + vbQuestion, "Confirm Play Time")

If strConfirm = vbNo Then
    WScript.Quit
End If

Do While Now() < dtDateTime
    WScript.Sleep 1000
Loop

objShell.Run """C:\Program Files\Windows Media Player\wmplayer.exe"" """ & strMP3Path & """", 1, False
