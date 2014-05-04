Option Explicit
Dim strOSArchitecture_Windir
Dim objFSO        	: Set objFSO        = WScript.CreateObject("Scripting.FileSystemObject")
Dim objShell      	: Set objShell      = CreateObject("Wscript.Shell")
Dim strWindir     	: strWinDir         = objShell.ExpandEnvironmentStrings("%WinDir%")
Dim strUrl, getOSVersion, version, OSVer, shell
Public strDestPath, overwrite
strDestPath= "C:\test\pshell3.msu"
overwrite = True
If objFSO.FolderExists(strWindir & "\syswow64") Then
     strOSArchitecture_Windir = "x64"
      Else
     strOSArchitecture_Windir = "x86"
  End If
'wscript.echo strOSArchitecture_Windir  
Set shell = CreateObject ("Wscript.Shell")
Set getOSVersion = shell.exec("%comspec% /c ver")
version = getOSVersion.stdout.readall
Select Case True
	Case InStr(version, "n 6.0") > 1 : OSVer = "6.0"
	Case InStr(version, "n 6.1") > 1 : OSVer = "6.1"
	Case Else : OSVer = "Unknown"
End Select	

If OSVer = "6.0" AND strOSArchitecture_Windir = "x64" Then
strUrl = "http://download.microsoft.com/download/E/7/6/E76850B8-DA6E-4FF5-8CCE-A24FC513FD16/Windows6.0-KB2506146-x64.msu"
ElseIf OSVer = "6.0" AND strOSArchitecture_Windir = "x86" Then
strUrl = "http://download.microsoft.com/download/E/7/6/E76850B8-DA6E-4FF5-8CCE-A24FC513FD16/Windows6.0-KB2506146-x86.msu"
ElseIf OSVer = "6.1" AND strOSArchitecture_Windir = "x64" Then
strUrl = "http://download.microsoft.com/download/E/7/6/E76850B8-DA6E-4FF5-8CCE-A24FC513FD16/Windows6.1-KB2506143-x64.msu"
ElseIf OSVer = "6.1" AND strOSArchitecture_Windir = "x86" Then
strUrl = "http://download.microsoft.com/download/E/7/6/E76850B8-DA6E-4FF5-8CCE-A24FC513FD16/Windows6.1-KB2506143-x86.msu"
	End If
Call Download (strUrl, strDestPath, overwrite)



Function Download ( ByRef strUrl, ByRef strDestPath, ByRef overwrite )
    Dim intStatusCode, objXMLHTTP, objADOStream, objFSO
    Set objFSO = CreateObject("Scripting.FileSystemObject")

    ' if the file exists already, and we're not overwriting, quit now
    If Not overwrite And objFSO.FileExists(strDestPath) Then
       ' WScript.Echo "Already exists - " & strDestPath
        Download = True
        Exit Function
    End If

   ' WScript.Echo "Downloading " & strUrl & " to " & strDestPath

    ' Fetch the file
    ' need to use ServerXMLHTTP so can set timeouts for downloading large files
    Set objXMLHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    objXMLHTTP.open "GET", strUrl, false
    objXMLHTTP.setTimeouts 1000 * 60 * 1, 1000 * 60 * 1, 1000 * 60 * 1, 1000 * 60 * 7
    objXMLHTTP.send()

    intStatusCode = objXMLHTTP.Status

    If intStatusCode = 200 Then
        Set objADOStream = CreateObject("ADODB.Stream")
        objADOStream.Open
        objADOStream.Type = 1 'adTypeBinary
        objADOStream.Write objXMLHTTP.ResponseBody
        objADOStream.Position = 0    'Set the stream position to the start

        'If the file already exists, delete it.
        'Otherwise, place the file in the specified location
        If objFSO.FileExists(strDestPath) Then objFSO.DeleteFile strDestPath

        objADOStream.SaveToFile strDestPath
        objADOStream.Close

        Set objADOStream = Nothing
    End If

    Set objXMLHTTP = Nothing
    Set objFSO = Nothing

    'WScript.Echo "Status code: " & intStatusCode & VBNewLine 

    If intStatusCode = 200 Then
        Download = True
    Else
        Download = False
    End If
End Function
